/**
 * Tracker Tool — Outlook task pane
 *
 * Reads the current mail item via Office.js, autofills via parse-email
 * (same AI flow as the web app), and POSTs to create-task with
 * client_id / client_name (required by the API).
 *
 * Auth: Supabase Auth REST password grant → store access_token +
 * refresh_token in OfficeRuntime.storage / RoamingSettings / localStorage.
 * Expired access tokens are refreshed via grant_type=refresh_token before
 * parse-email / create-task. Refresh failure clears session and prompts re-sign-in.
 * Uses anon key only — never service_role.
 */
(function () {
  "use strict";

  var Config = window.TrackerConfig;
  var emailContext = {
    subject: "",
    from: "",
    to: "",
    body: "",
    date: null,
    conversationId: null,
    internetMessageId: null,
  };

  /** Cached from REST after sign-in: [{id, name}, ...] */
  var clientsCache = [];
  var requestersCache = [];
  /** Matched client id from AI / datalist (optional) */
  var selectedClientId = null;
  var selectedRequesterId = null;
  var parseInFlight = false;

  // ---------------------------------------------------------------------------
  // Storage helpers
  // Prefer OfficeRuntime.storage, then RoamingSettings, then localStorage.
  // ---------------------------------------------------------------------------

  function roamingSettingsAvailable() {
    try {
      return (
        typeof Office !== "undefined" &&
        Office.context &&
        Office.context.roamingSettings &&
        typeof Office.context.roamingSettings.set === "function"
      );
    } catch (e) {
      return false;
    }
  }

  function roamingSet(key, value) {
    if (!roamingSettingsAvailable()) return;
    try {
      Office.context.roamingSettings.set(key, value);
      if (typeof Office.context.roamingSettings.saveAsync === "function") {
        Office.context.roamingSettings.saveAsync(function () {});
      }
    } catch (e) {}
  }

  function roamingGet(key) {
    if (!roamingSettingsAvailable()) return null;
    try {
      var v = Office.context.roamingSettings.get(key);
      return v != null && v !== "" ? String(v) : null;
    } catch (e) {
      return null;
    }
  }

  function roamingRemove(key) {
    if (!roamingSettingsAvailable()) return;
    try {
      Office.context.roamingSettings.remove(key);
      if (typeof Office.context.roamingSettings.saveAsync === "function") {
        Office.context.roamingSettings.saveAsync(function () {});
      }
    } catch (e) {}
  }

  function localStorageFallback(key) {
    try {
      return localStorage.getItem(key);
    } catch (e) {
      return null;
    }
  }

  function storageSet(key, value) {
    return new Promise(function (resolve) {
      roamingSet(key, value);
      try {
        localStorage.setItem(key, value);
      } catch (e0) {}
      try {
        if (
          typeof OfficeRuntime !== "undefined" &&
          OfficeRuntime.storage &&
          typeof OfficeRuntime.storage.setItem === "function"
        ) {
          OfficeRuntime.storage
            .setItem(key, value)
            .then(function () {
              resolve(true);
            })
            .catch(function () {
              resolve(true);
            });
          return;
        }
      } catch (e) {}
      resolve(true);
    });
  }

  function storageGet(key) {
    return new Promise(function (resolve) {
      try {
        if (
          typeof OfficeRuntime !== "undefined" &&
          OfficeRuntime.storage &&
          typeof OfficeRuntime.storage.getItem === "function"
        ) {
          OfficeRuntime.storage
            .getItem(key)
            .then(function (v) {
              if (v != null && v !== "") {
                resolve(v);
                return;
              }
              var roam = roamingGet(key);
              if (roam) {
                resolve(roam);
                return;
              }
              resolve(localStorageFallback(key));
            })
            .catch(function () {
              var roam = roamingGet(key);
              resolve(roam || localStorageFallback(key));
            });
          return;
        }
      } catch (e) {}
      var roam2 = roamingGet(key);
      resolve(roam2 || localStorageFallback(key));
    });
  }

  function storageRemove(key) {
    return new Promise(function (resolve) {
      roamingRemove(key);
      try {
        localStorage.removeItem(key);
      } catch (e0) {}
      try {
        if (
          typeof OfficeRuntime !== "undefined" &&
          OfficeRuntime.storage &&
          typeof OfficeRuntime.storage.removeItem === "function"
        ) {
          OfficeRuntime.storage
            .removeItem(key)
            .then(function () {
              resolve(true);
            })
            .catch(function () {
              resolve(true);
            });
          return;
        }
      } catch (e) {}
      resolve(true);
    });
  }

  // ---------------------------------------------------------------------------
  // UI helpers
  // ---------------------------------------------------------------------------

  function $(id) {
    return document.getElementById(id);
  }

  function showMsg(el, text, kind) {
    if (!el) return;
    el.textContent = text || "";
    el.className = "msg visible " + (kind || "info");
  }

  function clearMsg(el) {
    if (!el) return;
    el.textContent = "";
    el.className = "msg";
  }

  function setSignedInUi(email) {
    var auth = $("authPanel");
    var main = $("mainPanel");
    var form = $("formPanel");
    var chip = $("userChip");
    if (auth) auth.classList.add("hidden");
    if (main) main.classList.remove("hidden");
    if (form) form.classList.remove("hidden");
    if (chip) {
      chip.textContent = email || "Signed in";
      chip.title = email || "";
    }
  }

  function setSignedOutUi() {
    var auth = $("authPanel");
    var main = $("mainPanel");
    var form = $("formPanel");
    var chip = $("userChip");
    if (auth) auth.classList.remove("hidden");
    if (main) main.classList.add("hidden");
    if (form) form.classList.add("hidden");
    if (chip) {
      chip.textContent = "";
      chip.title = "";
    }
    clientsCache = [];
    requestersCache = [];
    selectedClientId = null;
    selectedRequesterId = null;
  }

  function fillDatalist(listEl, items) {
    if (!listEl) return;
    listEl.innerHTML = "";
    (items || []).forEach(function (item) {
      var opt = document.createElement("option");
      opt.value = item.name;
      listEl.appendChild(opt);
    });
  }

  function findClientByName(name) {
    if (!name) return null;
    var n = String(name).trim().toLowerCase();
    for (var i = 0; i < clientsCache.length; i++) {
      if (clientsCache[i].name.toLowerCase() === n) return clientsCache[i];
    }
    return null;
  }

  function findRequesterByName(name) {
    if (!name) return null;
    var emailName = String(name).trim().toLowerCase();
    var exact = null;
    for (var i = 0; i < requestersCache.length; i++) {
      if (requestersCache[i].name.toLowerCase() === emailName) {
        exact = requestersCache[i];
        break;
      }
    }
    if (exact) return exact;
    for (var j = 0; j < requestersCache.length; j++) {
      var r = requestersCache[j];
      var rn = r.name.toLowerCase();
      var rWords = rn.split(/\s+/);
      var eWords = emailName.split(/\s+/);
      var match =
        rWords.some(function (w) {
          return w.length > 2 && emailName.indexOf(w) !== -1;
        }) ||
        eWords.some(function (w) {
          return w.length > 2 && rn.indexOf(w) !== -1;
        });
      if (match) return r;
    }
    return null;
  }

  // ---------------------------------------------------------------------------
  // Auth (Supabase password grant + refresh_token renewal)
  // ---------------------------------------------------------------------------

  var SESSION_EXPIRED_MSG =
    "Your session expired. Please sign in again.";

  function mergeObjects(a, b) {
    var out = {};
    var k;
    if (a) {
      for (k in a) {
        if (Object.prototype.hasOwnProperty.call(a, k)) out[k] = a[k];
      }
    }
    if (b) {
      for (k in b) {
        if (Object.prototype.hasOwnProperty.call(b, k)) out[k] = b[k];
      }
    }
    return out;
  }

  function persistSession(data, email) {
    var chain = storageSet(Config.storageTokenKey, data.access_token);
    if (data.refresh_token) {
      chain = chain.then(function () {
        return storageSet(Config.storageRefreshTokenKey, data.refresh_token);
      });
    }
    if (email) {
      chain = chain.then(function () {
        return storageSet(Config.storageUserKey, email);
      });
    }
    return chain.then(function () {
      return data;
    });
  }

  function clearStoredSession() {
    return storageRemove(Config.storageTokenKey)
      .then(function () {
        return storageRemove(Config.storageRefreshTokenKey);
      })
      .then(function () {
        return storageRemove(Config.storageUserKey);
      });
  }

  function signIn(email, password) {
    if (Config.isPlaceholder()) {
      return Promise.reject(
        new Error(
          "Config still has PLACEHOLDER_* values. Set supabaseUrl and supabaseAnonKey in src/config.js."
        )
      );
    }

    var url = Config.authTokenUrl();
    return fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        apikey: Config.supabaseAnonKey,
      },
      body: JSON.stringify({ email: email, password: password }),
    }).then(function (res) {
      return res.json().then(function (data) {
        if (!res.ok) {
          var msg =
            (data && (data.error_description || data.msg || data.error)) ||
            "Sign-in failed (" + res.status + ")";
          throw new Error(msg);
        }
        if (!data || !data.access_token) {
          throw new Error("No access_token in auth response");
        }
        return persistSession(data, email);
      });
    });
  }

  function signOut() {
    return clearStoredSession();
  }

  function forceReSignIn(message) {
    return clearStoredSession().then(function () {
      setSignedOutUi();
      var msg = message || SESSION_EXPIRED_MSG;
      showMsg($("authMsg"), msg, "error");
      clearMsg($("aiMsg"));
      clearMsg($("taskMsg"));
      return null;
    });
  }

  /**
   * Decode JWT payload without verifying signature (client-side expiry check only).
   */
  function decodeJwtPayload(token) {
    if (!token || typeof token !== "string") return null;
    var parts = token.split(".");
    if (parts.length < 2) return null;
    try {
      var b64 = parts[1].replace(/-/g, "+").replace(/_/g, "/");
      while (b64.length % 4) b64 += "=";
      var json = null;
      if (typeof atob === "function") {
        json = atob(b64);
      } else if (typeof Buffer !== "undefined") {
        json = Buffer.from(b64, "base64").toString("utf8");
      }
      return json ? JSON.parse(json) : null;
    } catch (e) {
      return null;
    }
  }

  function isAccessTokenExpired(token, skewSeconds) {
    var skew = typeof skewSeconds === "number" ? skewSeconds : 60;
    var payload = decodeJwtPayload(token);
    if (!payload || typeof payload.exp !== "number") {
      // Unknown shape — treat as needing refresh if we have a refresh_token
      return true;
    }
    var now = Math.floor(Date.now() / 1000);
    return payload.exp <= now + skew;
  }

  function isAuthFailureMessage(msg) {
    if (!msg) return false;
    var s = String(msg).toLowerCase();
    return (
      s.indexOf("jwt expired") !== -1 ||
      (s.indexOf("jwt") !== -1 && s.indexOf("expir") !== -1) ||
      s.indexOf("invalid jwt") !== -1 ||
      s.indexOf("not authenticated") !== -1 ||
      s.indexOf("invalid claim") !== -1 ||
      s.indexOf("session expired") !== -1 ||
      s.indexOf("token is expired") !== -1
    );
  }

  /**
   * POST /auth/v1/token?grant_type=refresh_token
   * Persists rotated access_token + refresh_token on success.
   */
  function refreshSession() {
    return storageGet(Config.storageRefreshTokenKey).then(function (refreshToken) {
      if (!refreshToken) {
        throw new Error("No refresh_token — please sign in again.");
      }
      return fetch(Config.authRefreshUrl(), {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          apikey: Config.supabaseAnonKey,
        },
        body: JSON.stringify({ refresh_token: refreshToken }),
      }).then(function (res) {
        return res.json().then(function (data) {
          if (!res.ok) {
            var msg =
              (data && (data.error_description || data.msg || data.error)) ||
              "Token refresh failed (" + res.status + ")";
            var err = new Error(
              typeof msg === "string" ? msg : JSON.stringify(msg)
            );
            err.status = res.status;
            err.refreshFailed = true;
            throw err;
          }
          if (!data || !data.access_token) {
            var err2 = new Error("No access_token in refresh response");
            err2.refreshFailed = true;
            throw err2;
          }
          return storageGet(Config.storageUserKey).then(function (email) {
            return persistSession(data, email || null).then(function () {
              return data.access_token;
            });
          });
        });
      });
    });
  }

  /**
   * Return a usable access_token, refreshing via refresh_token when expired.
   * On refresh failure: clear session and show re-sign-in UI (rejects with SESSION_EXPIRED_MSG).
   */
  function ensureValidAccessToken() {
    return storageGet(Config.storageTokenKey).then(function (token) {
      if (token && !isAccessTokenExpired(token)) {
        return token;
      }
      // Missing or expired — try refresh_token
      return refreshSession().catch(function (err) {
        return forceReSignIn(SESSION_EXPIRED_MSG).then(function () {
          var e = new Error(SESSION_EXPIRED_MSG);
          e.authRequired = true;
          e.cause = err;
          throw e;
        });
      });
    });
  }

  function authHeaders(token) {
    return {
      "Content-Type": "application/json",
      Authorization: "Bearer " + token,
      apikey: Config.supabaseAnonKey,
    };
  }

  /**
   * Run an authenticated fetch. On 401 / JWT-expired body, refresh once and retry.
   * If refresh fails, clear session and reject with a clear re-sign-in message.
   */
  function fetchWithAuth(url, options) {
    options = options || {};
    return ensureValidAccessToken().then(function (token) {
      var headers = mergeObjects(options.headers || {}, authHeaders(token));
      return fetch(url, mergeObjects(options, { headers: headers })).then(
        function (res) {
          if (res.status !== 401) return res;
          // Try one refresh + retry
          return res
            .text()
            .then(function (text) {
              var data = null;
              try {
                data = text ? JSON.parse(text) : null;
              } catch (e) {
                data = null;
              }
              var errMsg =
                (data && (data.error || data.message || data.msg)) ||
                text ||
                "Unauthorized";
              return String(errMsg);
            })
            .then(function (errMsg) {
              if (!isAuthFailureMessage(errMsg) && errMsg.indexOf("401") === -1) {
                // Still retry refresh on any 401 from our APIs
              }
              return refreshSession()
                .then(function (newToken) {
                  var headers2 = mergeObjects(
                    options.headers || {},
                    authHeaders(newToken)
                  );
                  return fetch(
                    url,
                    mergeObjects(options, { headers: headers2 })
                  );
                })
                .catch(function () {
                  return forceReSignIn(SESSION_EXPIRED_MSG).then(function () {
                    var e = new Error(SESSION_EXPIRED_MSG);
                    e.authRequired = true;
                    throw e;
                  });
                });
            });
        }
      );
    });
  }

  // ---------------------------------------------------------------------------
  // Load clients + requesters (user JWT + RLS)
  // ---------------------------------------------------------------------------

  function fetchClientsAndRequesters() {
    var clientsUrl = Config.restUrl(
      "clients",
      "select=id,name&order=name.asc"
    );
    var requestersUrl = Config.restUrl(
      "requesters",
      "select=id,name&order=name.asc"
    );

    function parseList(res, label) {
      return res.text().then(function (text) {
        var data = null;
        try {
          data = text ? JSON.parse(text) : null;
        } catch (e) {
          data = null;
        }
        if (!res.ok) {
          var msg =
            (data && (data.message || data.error || data.msg)) ||
            text ||
            "Failed to load " + label + " (" + res.status + ")";
          var err = new Error(
            typeof msg === "string" ? msg : JSON.stringify(msg)
          );
          err.status = res.status;
          throw err;
        }
        return Array.isArray(data) ? data : [];
      });
    }

    return Promise.all([
      fetchWithAuth(clientsUrl, {
        method: "GET",
        headers: { Accept: "application/json" },
      }).then(function (res) {
        return parseList(res, "clients");
      }),
      fetchWithAuth(requestersUrl, {
        method: "GET",
        headers: { Accept: "application/json" },
      }).then(function (res) {
        return parseList(res, "requesters");
      }),
    ]).then(function (pair) {
      clientsCache = pair[0];
      requestersCache = pair[1];
      fillDatalist($("clientList"), clientsCache);
      fillDatalist($("requesterList"), requestersCache);
      return { clients: clientsCache, requesters: requestersCache };
    });
  }

  // ---------------------------------------------------------------------------
  // Office.js — read current mail item
  // ---------------------------------------------------------------------------

  function formatRecipient(rec) {
    if (!rec) return "";
    if (typeof rec === "string") return rec;
    var name = rec.displayName || "";
    var addr = rec.emailAddress || "";
    if (name && addr) return name + " <" + addr + ">";
    return name || addr || "";
  }

  function formatRecipients(list) {
    if (!list || !list.length) return "";
    return list
      .map(function (r) {
        return formatRecipient(r);
      })
      .join(", ");
  }

  function loadMailItem() {
    return new Promise(function (resolve, reject) {
      if (typeof Office === "undefined" || !Office.context || !Office.context.mailbox) {
        reject(new Error("Office mailbox context not available"));
        return;
      }

      var item = Office.context.mailbox.item;
      if (!item) {
        reject(new Error("No mail item selected"));
        return;
      }

      emailContext.subject = item.subject || "";
      emailContext.from = formatRecipient(item.from);
      emailContext.to = formatRecipients(item.to);
      emailContext.date = item.dateTimeCreated || null;
      emailContext.conversationId = item.conversationId || null;
      emailContext.internetMessageId = item.internetMessageId || null;

      if (item.body && typeof item.body.getAsync === "function") {
        item.body.getAsync(Office.CoercionType.Text, function (result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            emailContext.body = result.value || "";
          } else {
            emailContext.body =
              "(Could not read body: " +
              (result.error && result.error.message
                ? result.error.message
                : "unknown") +
              ")";
          }
          resolve(emailContext);
        });
      } else {
        emailContext.body = "";
        resolve(emailContext);
      }
    });
  }

  function applyEmailToUi(ctx) {
    $("previewFrom").textContent = ctx.from || "—";
    $("previewTo").textContent = ctx.to || "—";
    $("previewSubject").textContent = ctx.subject || "—";
    var dateStr = "—";
    if (ctx.date) {
      try {
        dateStr = new Date(ctx.date).toLocaleString();
      } catch (e) {
        dateStr = String(ctx.date);
      }
    }
    $("previewDate").textContent = dateStr;

    var preview = ctx.body || "";
    if (preview.length > 2000) {
      preview = preview.slice(0, 2000) + "\n…";
    }
    $("bodyPreview").textContent = preview || "(empty)";

    // Baseline premill from Outlook (AI may refine)
    $("fieldTitle").value = ctx.subject || "";
    $("fieldDescription").value = "";
    $("fieldRequester").value = ctx.from || "";
    $("fieldClient").value = "";
    selectedClientId = null;
    selectedRequesterId = null;
  }

  /**
   * Build RFC822-ish text for parse-email (same shape as .eml text path).
   */
  function buildEmailContentForParse(ctx) {
    var parts = [];
    if (ctx.from) parts.push("From: " + ctx.from);
    if (ctx.to) parts.push("To: " + ctx.to);
    if (ctx.subject) parts.push("Subject: " + ctx.subject);
    if (ctx.date) {
      try {
        parts.push("Date: " + new Date(ctx.date).toUTCString());
      } catch (e) {
        parts.push("Date: " + String(ctx.date));
      }
    }
    parts.push("");
    parts.push(ctx.body || "");
    return parts.join("\n");
  }

  // ---------------------------------------------------------------------------
  // parse-email AI autofill (matches web CreateTaskDialog flow)
  // ---------------------------------------------------------------------------

  function parseEmailAutofill() {
    if (parseInFlight) return Promise.resolve(null);
    parseInFlight = true;
    clearMsg($("aiMsg"));
    showMsg($("aiMsg"), "Parsing email with AI…", "info");

    return ensureValidAccessToken()
      .then(function () {
        var ensureLists =
          clientsCache.length || requestersCache.length
            ? Promise.resolve()
            : fetchClientsAndRequesters();

        return ensureLists.then(function () {
          var payload = {
            emailContent: buildEmailContentForParse(emailContext),
            clients: clientsCache.map(function (c) {
              return c.name;
            }),
            requesters: requestersCache.map(function (r) {
              return r.name;
            }),
          };

          return fetchWithAuth(Config.resolvedParseEmailUrl(), {
            method: "POST",
            body: JSON.stringify(payload),
          }).then(function (res) {
            return res.text().then(function (text) {
              var data = null;
              try {
                data = text ? JSON.parse(text) : null;
              } catch (e) {
                data = null;
              }
              if (!res.ok) {
                var errMsg =
                  (data && (data.error || data.message)) ||
                  text ||
                  "parse-email failed (" + res.status + ")";
                var err = new Error(
                  typeof errMsg === "string" ? errMsg : JSON.stringify(errMsg)
                );
                err.status = res.status;
                throw err;
              }
              return data;
            });
          });
        });
      })
      .then(function (data) {
        if (!data) return null;

        var hasAny =
          data.title ||
          data.description ||
          data.client ||
          data.requester ||
          data.emailFrom;
        if (!hasAny) {
          showMsg(
            $("aiMsg"),
            "AI returned no fields — fill title and client manually.",
            "info"
          );
          return data;
        }

        if (data.title) $("fieldTitle").value = data.title;
        if (data.description) $("fieldDescription").value = data.description;

        if (data.emailFrom) {
          emailContext.from = data.emailFrom;
          $("previewFrom").textContent = data.emailFrom;
        }
        if (data.emailTo) {
          emailContext.to = data.emailTo;
          $("previewTo").textContent = data.emailTo;
        }
        if (data.emailDate) {
          try {
            emailContext.date = new Date(data.emailDate);
            $("previewDate").textContent = emailContext.date.toLocaleString();
          } catch (e) {}
        }

        selectedClientId = null;
        if (data.client) {
          var matchedClient = findClientByName(data.client);
          if (matchedClient) {
            $("fieldClient").value = matchedClient.name;
            selectedClientId = matchedClient.id;
          } else {
            // Show AI suggestion; create-task will fail if not exact match
            $("fieldClient").value = data.client;
          }
        }

        selectedRequesterId = null;
        if (data.requester) {
          var matchedReq = findRequesterByName(data.requester);
          if (matchedReq) {
            $("fieldRequester").value = matchedReq.name;
            selectedRequesterId = matchedReq.id;
          } else {
            $("fieldRequester").value = data.requester;
          }
        } else if (!$("fieldRequester").value && emailContext.from) {
          $("fieldRequester").value = emailContext.from;
        }

        showMsg($("aiMsg"), "Email parsed — review client before creating.", "success");
        return data;
      })
      .catch(function (err) {
        if (err && err.authRequired) {
          // forceReSignIn already updated auth UI
          showMsg(
            $("aiMsg"),
            SESSION_EXPIRED_MSG + " Sign in to enable AI parse.",
            "error"
          );
          return null;
        }
        var raw = err && err.message ? err.message : String(err);
        if (isAuthFailureMessage(raw)) {
          return forceReSignIn(SESSION_EXPIRED_MSG).then(function () {
            showMsg(
              $("aiMsg"),
              SESSION_EXPIRED_MSG + " Sign in to enable AI parse.",
              "error"
            );
            return null;
          });
        }
        showMsg(
          $("aiMsg"),
          "AI parse skipped: " + raw + ". Fill fields manually.",
          "error"
        );
        return null;
      })
      .then(function (result) {
        parseInFlight = false;
        return result;
      });
  }

  function afterMailLoaded() {
    applyEmailToUi(emailContext);
    return ensureValidAccessToken()
      .then(function () {
        return fetchClientsAndRequesters()
          .catch(function (err) {
            if (err && err.authRequired) return null;
            var raw = err && err.message ? err.message : String(err);
            if (isAuthFailureMessage(raw)) {
              return forceReSignIn(SESSION_EXPIRED_MSG);
            }
            showMsg(
              $("aiMsg"),
              "Could not load clients: " + raw,
              "error"
            );
          })
          .then(function () {
            // Only parse when still signed in
            return storageGet(Config.storageTokenKey).then(function (token) {
              if (!token) return null;
              return parseEmailAutofill();
            });
          });
      })
      .catch(function (err) {
        if (err && err.authRequired) return null;
        return null;
      });
  }

  // ---------------------------------------------------------------------------
  // Create Task — API needs client_id OR client_name (not "client")
  // ---------------------------------------------------------------------------

  function resolveClientForSubmit() {
    var name = ($("fieldClient").value || "").trim();
    if (!name) {
      return {
        error:
          "Client is required. Pick an existing Tracker client (exact name match).",
      };
    }
    var matched = findClientByName(name);
    if (matched) {
      return { client_id: matched.id, client_name: matched.name };
    }
    // Allow sending client_name — API does case-insensitive exact match
    return { client_name: name };
  }

  function createTask() {
    return ensureValidAccessToken().then(function () {
      var title = ($("fieldTitle").value || "").trim();
      if (!title) {
        throw new Error("Title is required");
      }

      var clientResolved = resolveClientForSubmit();
      if (clientResolved.error) {
        throw new Error(clientResolved.error);
      }

      var requesterRaw = ($("fieldRequester").value || "").trim();
      var matchedReq = requesterRaw ? findRequesterByName(requesterRaw) : null;

      var payload = {
        title: title,
        description: ($("fieldDescription").value || "").trim() || null,
        email_from: emailContext.from || null,
        email_to: emailContext.to || null,
        email_date: emailContext.date
          ? new Date(emailContext.date).toISOString()
          : null,
      };

      if (clientResolved.client_id) {
        payload.client_id = clientResolved.client_id;
      }
      if (clientResolved.client_name) {
        payload.client_name = clientResolved.client_name;
      }

      if (matchedReq) {
        payload.requester_id = matchedReq.id;
      } else if (requesterRaw) {
        payload.requester_name = requesterRaw;
      }

      // Omit status_id → create-task defaults to "To Do" (same as web)

      var url = Config.resolvedCreateTaskUrl();

      return fetchWithAuth(url, {
        method: "POST",
        body: JSON.stringify(payload),
      }).then(function (res) {
        return res.text().then(function (text) {
          var data = null;
          try {
            data = text ? JSON.parse(text) : null;
          } catch (e) {
            data = { raw: text };
          }
          if (!res.ok) {
            var errMsg =
              (data && (data.error || data.message || data.msg)) ||
              text ||
              "Create task failed (" + res.status + ")";
            if (
              typeof errMsg === "string" &&
              /client_id or client_name/i.test(errMsg)
            ) {
              errMsg =
                "Client is required and must match an existing Tracker client exactly.";
            }
            if (
              typeof errMsg === "string" &&
              /Client .* not found/i.test(errMsg)
            ) {
              errMsg =
                errMsg +
                " Pick a client from the list (exact name). The add-in cannot create new clients.";
            }
            if (res.status === 404 || res.status === 0) {
              errMsg +=
                " — check that Tracker Tool is available and try again.";
            }
            if (typeof errMsg === "string" && isAuthFailureMessage(errMsg)) {
              errMsg = SESSION_EXPIRED_MSG;
            }
            var err = new Error(
              typeof errMsg === "string" ? errMsg : JSON.stringify(errMsg)
            );
            err.status = res.status;
            err.data = data;
            if (isAuthFailureMessage(String(errMsg))) {
              err.authRequired = true;
            }
            throw err;
          }
          return data;
        });
      });
    });
  }

  // ---------------------------------------------------------------------------
  // Boot
  // ---------------------------------------------------------------------------

  function bindEvents() {
    $("btnSignIn").addEventListener("click", function () {
      clearMsg($("authMsg"));
      var email = ($("authEmail").value || "").trim();
      var password = $("authPassword").value || "";
      if (!email || !password) {
        showMsg($("authMsg"), "Enter email and password.", "error");
        return;
      }
      $("btnSignIn").disabled = true;
      signIn(email, password)
        .then(function () {
          showMsg($("authMsg"), "Signed in.", "success");
          setSignedInUi(email);
          return loadMailItem().then(function () {
            return afterMailLoaded();
          });
        })
        .catch(function (err) {
          showMsg(
            $("authMsg"),
            err && err.message ? err.message : String(err),
            "error"
          );
        })
        .then(function () {
          $("btnSignIn").disabled = false;
        });
    });

    $("btnSignOut").addEventListener("click", function () {
      signOut().then(function () {
        setSignedOutUi();
        clearMsg($("taskMsg"));
        clearMsg($("authMsg"));
        clearMsg($("aiMsg"));
      });
    });

    $("btnRefresh").addEventListener("click", function () {
      clearMsg($("taskMsg"));
      clearMsg($("aiMsg"));
      loadMailItem()
        .then(function () {
          return afterMailLoaded();
        })
        .catch(function (err) {
          showMsg(
            $("taskMsg"),
            err && err.message ? err.message : String(err),
            "error"
          );
        });
    });

    $("btnCreateTask").addEventListener("click", function () {
      clearMsg($("taskMsg"));
      // Sync client id if user typed/picked from datalist
      var typed = ($("fieldClient").value || "").trim();
      var m = findClientByName(typed);
      selectedClientId = m ? m.id : null;

      $("btnCreateTask").disabled = true;
      createTask()
        .then(function (data) {
          var taskId =
            (data && data.task && data.task.id) ||
            (data && data.id) ||
            "";
          var ok =
            "Task created." +
            (taskId ? " ID: " + taskId : "") +
            (Config.trackerUrl
              ? " Open Tracker: " + Config.trackerUrl
              : "");
          showMsg($("taskMsg"), ok, "success");
        })
        .catch(function (err) {
          if (err && err.authRequired) {
            showMsg($("taskMsg"), SESSION_EXPIRED_MSG, "error");
            return;
          }
          var raw = err && err.message ? err.message : String(err);
          if (isAuthFailureMessage(raw)) {
            forceReSignIn(SESSION_EXPIRED_MSG);
            showMsg($("taskMsg"), SESSION_EXPIRED_MSG, "error");
            return;
          }
          showMsg($("taskMsg"), raw, "error");
        })
        .then(function () {
          $("btnCreateTask").disabled = false;
        });
    });

    // Keep selectedClientId in sync when client field changes
    var clientInput = $("fieldClient");
    if (clientInput) {
      clientInput.addEventListener("change", function () {
        var m = findClientByName(clientInput.value);
        selectedClientId = m ? m.id : null;
      });
      clientInput.addEventListener("input", function () {
        var m = findClientByName(clientInput.value);
        selectedClientId = m ? m.id : null;
      });
    }
  }

  function boot() {
    if (Config.isPlaceholder()) {
      $("configWarn").classList.remove("hidden");
    }

    bindEvents();

    storageGet(Config.storageTokenKey).then(function (token) {
      return storageGet(Config.storageRefreshTokenKey).then(function (refresh) {
        return storageGet(Config.storageUserKey).then(function (email) {
          if (!token && !refresh) {
            setSignedOutUi();
            return;
          }
          // Refresh access token if needed, then continue
          ensureValidAccessToken()
            .then(function () {
              setSignedInUi(email);
              return loadMailItem().then(function () {
                return afterMailLoaded();
              });
            })
            .catch(function (err) {
              if (err && err.authRequired) return;
              showMsg(
                $("taskMsg"),
                err && err.message ? err.message : String(err),
                "error"
              );
            });
        });
      });
    });
  }

  window.TrackerTaskpane = {
    boot: boot,
    loadMailItem: loadMailItem,
    parseEmailAutofill: parseEmailAutofill,
  };

  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(function () {
      boot();
    });
  } else {
    document.addEventListener("DOMContentLoaded", function () {
      if (typeof Office !== "undefined" && Office.onReady) {
        Office.onReady(function () {
          boot();
        });
      } else {
        boot();
      }
    });
  }
})();
