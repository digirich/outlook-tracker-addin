/**
 * Tracker Tool — Outlook task pane
 *
 * Reads the current mail item via Office.js and posts a create-task stub
 * to the planned Supabase edge function (Bearer user JWT).
 *
 * Auth: Supabase Auth REST password grant → store access_token in
 * OfficeRuntime.storage (fallback: localStorage).
 *
 * NOTE: create-task API is forthcoming. This client is ready to call it.
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

  // ---------------------------------------------------------------------------
  // Storage helpers (OfficeRuntime.storage preferred in Office hosts)
  // ---------------------------------------------------------------------------

  function storageSet(key, value) {
    return new Promise(function (resolve) {
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
              try {
                localStorage.setItem(key, value);
              } catch (e) {}
              resolve(true);
            });
          return;
        }
      } catch (e) {}
      try {
        localStorage.setItem(key, value);
      } catch (e2) {}
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
              } else {
                resolve(localStorageFallback(key));
              }
            })
            .catch(function () {
              resolve(localStorageFallback(key));
            });
          return;
        }
      } catch (e) {}
      resolve(localStorageFallback(key));
    });
  }

  function storageRemove(key) {
    return new Promise(function (resolve) {
      try {
        if (
          typeof OfficeRuntime !== "undefined" &&
          OfficeRuntime.storage &&
          typeof OfficeRuntime.storage.removeItem === "function"
        ) {
          OfficeRuntime.storage
            .removeItem(key)
            .then(function () {
              try {
                localStorage.removeItem(key);
              } catch (e) {}
              resolve(true);
            })
            .catch(function () {
              try {
                localStorage.removeItem(key);
              } catch (e2) {}
              resolve(true);
            });
          return;
        }
      } catch (e) {}
      try {
        localStorage.removeItem(key);
      } catch (e3) {}
      resolve(true);
    });
  }

  function localStorageFallback(key) {
    try {
      return localStorage.getItem(key);
    } catch (e) {
      return null;
    }
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
  }

  // ---------------------------------------------------------------------------
  // Auth (Supabase password grant)
  // ---------------------------------------------------------------------------

  /**
   * POST /auth/v1/token?grant_type=password
   * Headers: apikey, Content-Type
   * Body: { email, password }
   *
   * Stores access_token for Bearer calls to create-task (when that API exists).
   */
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
        return storageSet(Config.storageTokenKey, data.access_token).then(
          function () {
            return storageSet(Config.storageUserKey, email).then(function () {
              return data;
            });
          }
        );
      });
    });
  }

  function signOut() {
    return storageRemove(Config.storageTokenKey).then(function () {
      return storageRemove(Config.storageUserKey);
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

      // Body requires getAsync
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

    $("fieldTitle").value = ctx.subject || "";
    $("fieldDescription").value = preview || "";
    $("fieldRequester").value = ctx.from || "";
  }

  // ---------------------------------------------------------------------------
  // Create Task (stub POST — endpoint not built yet)
  // ---------------------------------------------------------------------------

  /**
   * Planned payload for POST /functions/v1/create-task
   * Authorization: Bearer <user JWT>
   * apikey: anon key
   *
   * Fields: title, description, email_from, email_to, email_date
   * Optional later: client_id, status, requester, conversation_id, internet_message_id
   */
  function createTask() {
    return storageGet(Config.storageTokenKey).then(function (token) {
      if (!token) {
        throw new Error("Not signed in");
      }

      var payload = {
        title: ($("fieldTitle").value || "").trim(),
        description: ($("fieldDescription").value || "").trim(),
        email_from: emailContext.from || "",
        email_to: emailContext.to || "",
        email_date: emailContext.date
          ? new Date(emailContext.date).toISOString()
          : null,
        // Optional / forthcoming
        client: ($("fieldClient").value || "").trim() || null,
        // client_id: null, // TODO when clients list is wired
        requester: ($("fieldRequester").value || "").trim() || null,
        status: $("fieldStatus").value || "open",
        conversation_id: emailContext.conversationId || null,
        internet_message_id: emailContext.internetMessageId || null,
      };

      if (!payload.title) {
        throw new Error("Title is required");
      }

      var url = Config.resolvedCreateTaskUrl();

      if (Config.isPlaceholder()) {
        // Still attempt the call so the shape can be verified once URLs are set;
        // surface a clear message if placeholders remain.
        console.warn(
          "[Tracker Tool] create-task URL may still be a placeholder:",
          url
        );
      }

      return fetch(url, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: "Bearer " + token,
          apikey: Config.supabaseAnonKey,
        },
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
            // Helpful hint while endpoint is not built
            if (res.status === 404 || res.status === 0) {
              errMsg +=
                " — create-task edge function may not be deployed yet.";
            }
            var err = new Error(
              typeof errMsg === "string" ? errMsg : JSON.stringify(errMsg)
            );
            err.status = res.status;
            err.data = data;
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
          return loadMailItem().then(applyEmailToUi);
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
      });
    });

    $("btnRefresh").addEventListener("click", function () {
      clearMsg($("taskMsg"));
      loadMailItem()
        .then(applyEmailToUi)
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
      $("btnCreateTask").disabled = true;
      createTask()
        .then(function (data) {
          var ok =
            "Task created." +
            (data && data.id ? " ID: " + data.id : "") +
            (Config.trackerUrl
              ? " Open Tracker: " + Config.trackerUrl
              : "");
          showMsg($("taskMsg"), ok, "success");
        })
        .catch(function (err) {
          showMsg(
            $("taskMsg"),
            err && err.message ? err.message : String(err),
            "error"
          );
        })
        .then(function () {
          $("btnCreateTask").disabled = false;
        });
    });
  }

  function boot() {
    if (Config.isPlaceholder()) {
      $("configWarn").classList.remove("hidden");
    }

    bindEvents();

    storageGet(Config.storageTokenKey).then(function (token) {
      return storageGet(Config.storageUserKey).then(function (email) {
        if (token) {
          setSignedInUi(email);
          loadMailItem()
            .then(applyEmailToUi)
            .catch(function (err) {
              showMsg(
                $("taskMsg"),
                err && err.message ? err.message : String(err),
                "error"
              );
            });
        } else {
          setSignedOutUi();
        }
      });
    });
  }

  window.TrackerTaskpane = { boot: boot, loadMailItem: loadMailItem };

  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(function (info) {
      // info.host === Office.HostType.Outlook when running in Outlook
      boot();
    });
  } else {
    // Browser preview / late Office.js load
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
