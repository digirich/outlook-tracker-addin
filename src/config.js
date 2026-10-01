/**
 * Tracker Tool — configuration
 *
 * Uses the public Supabase anon key only (safe with RLS).
 * Never put the service_role key here.
 */
(function (global) {
  "use strict";

  var Config = {
    /** Display name shown in the task pane */
    appName: "Tracker Tool",

    /** Live tracker web app */
    trackerUrl: "https://tracker.seoandweb.co.uk",

    /** Supabase project URL */
    supabaseUrl: "https://zsabmwwtflsonkjnufxc.supabase.co",

    /**
     * Supabase anon (public) key — safe for client use with RLS.
     * Never put the service_role key here.
     */
    supabaseAnonKey: "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InpzYWJtd3d0Zmxzb25ram51ZnhjIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjY0MTEyMDIsImV4cCI6MjA4MTk4NzIwMn0.jSX4h45_FIzrNa7Bk1_lvBKdkCNbUGI3-L3VdGhAWck",

    /**
     * create-task edge function.
     * POST with Bearer user JWT + apikey (anon).
     * Required: title + client_id OR client_name
     * Optional: description, requester_id|requester_name, status_id,
     *           email_from, email_to, email_date, emailContent, emailBase64
     */
    createTaskUrl:
      "https://zsabmwwtflsonkjnufxc.supabase.co/functions/v1/create-task",

    /**
     * parse-email edge function (AI autofill).
     * POST with Bearer user JWT + apikey (anon).
     * Body: { emailContent|emailBase64, clients: string[], requesters: string[] }
     * Returns: { title, description, client, requester, emailFrom, emailTo, emailDate }
     */
    parseEmailUrl:
      "https://zsabmwwtflsonkjnufxc.supabase.co/functions/v1/parse-email",

    /** OfficeRuntime.storage / RoamingSettings / localStorage key for access token */
    storageTokenKey: "tracker_access_token",

    /** Same stores: refresh_token for silent JWT renewal (anon key only) */
    storageRefreshTokenKey: "tracker_refresh_token",

    /** Cached user email (display only) */
    storageUserKey: "tracker_user_email",

    /**
     * Returns true when config still has placeholders (dev guard).
     */
    isPlaceholder: function () {
      return (
        String(this.supabaseUrl).indexOf("PLACEHOLDER_") === 0 ||
        String(this.supabaseAnonKey).indexOf("PLACEHOLDER_") === 0
      );
    },

    /**
     * Build absolute Supabase Auth token URL (password grant).
     * POST /auth/v1/token?grant_type=password
     */
    authTokenUrl: function () {
      return (
        this.supabaseUrl.replace(/\/$/, "") +
        "/auth/v1/token?grant_type=password"
      );
    },

    /**
     * Supabase Auth refresh_token grant.
     * POST /auth/v1/token?grant_type=refresh_token
     * Body: { refresh_token }
     */
    authRefreshUrl: function () {
      return (
        this.supabaseUrl.replace(/\/$/, "") +
        "/auth/v1/token?grant_type=refresh_token"
      );
    },

    /** REST base for PostgREST (clients / requesters lists) */
    restUrl: function (table, query) {
      return (
        this.supabaseUrl.replace(/\/$/, "") +
        "/rest/v1/" +
        table +
        (query ? "?" + query : "")
      );
    },

    /**
     * Resolve the create-task URL, substituting supabaseUrl if templated.
     */
    resolvedCreateTaskUrl: function () {
      return String(this.createTaskUrl).replace(
        "PLACEHOLDER_SUPABASE_URL",
        this.supabaseUrl.replace(/\/$/, "")
      );
    },

    resolvedParseEmailUrl: function () {
      return String(this.parseEmailUrl).replace(
        "PLACEHOLDER_SUPABASE_URL",
        this.supabaseUrl.replace(/\/$/, "")
      );
    },
  };

  global.TrackerConfig = Config;
})(typeof window !== "undefined" ? window : this);
