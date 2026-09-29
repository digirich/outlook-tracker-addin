/**
 * Tracker Tool — configuration
 *
 * TODO: Replace every PLACEHOLDER_* value before production use.
 * Do not commit real secrets. Prefer injecting via build/env at deploy time.
 */
(function (global) {
  "use strict";

  var Config = {
    /** Display name shown in the task pane */
    appName: "Tracker Tool",

    /** Live tracker web app (reference only; add-in talks to Supabase) */
    trackerUrl: "https://tracker.seoandweb.co.uk",

    /**
     * TODO: Supabase project URL
     * Example: https://YOUR_PROJECT.supabase.co
     */
    supabaseUrl: "https://zsabmwwtflsonkjnufxc.supabase.co",

    /**
     * TODO: Supabase anon (public) key — safe for client use with RLS.
     * Never put the service_role key here.
     */
    supabaseAnonKey: "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InpzYWJtd3d0Zmxzb25ram51ZnhjIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjY0MTEyMDIsImV4cCI6MjA4MTk4NzIwMn0.jSX4h45_FIzrNa7Bk1_lvBKdkCNbUGI3-L3VdGhAWck",

    /**
     * TODO: Edge function that creates a task from email context.
     * Planned: POST with Bearer user JWT
     * Body: { title, description, email_from, email_to, email_date, client_id? }
     */
    createTaskUrl:
      "https://zsabmwwtflsonkjnufxc.supabase.co/functions/v1/create-task",

    /**
     * Existing edge function for email parsing (optional helper).
     * POST .../functions/v1/parse-email
     */
    parseEmailUrl:
      "https://zsabmwwtflsonkjnufxc.supabase.co/functions/v1/parse-email",

    /** localStorage / OfficeRuntime.storage key for the access token */
    storageTokenKey: "tracker_access_token",

    /** localStorage key for cached user email (display only) */
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
     * Resolve create-task URL, substituting supabaseUrl if still templated.
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

  // Known project host (non-secret) — still require anon key + confirm URLs.
  // Real values: https://zsabmwwtflsonkjnufxc.supabase.co
  // Leave PLACEHOLDER_* until you deliberately set them in this file or a build step.

  global.TrackerConfig = Config;
})(typeof window !== "undefined" ? window : this);
