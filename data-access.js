// data-access.js — User access configuration
// Add entries to restrict users to specific regions.
// Users NOT in this list can see all regions (unrestricted access).
// Region values must match exactly what appears in the "Region L2" filter
// e.g. "CS UKI", "CS Nordic", "CS France", "CS BeNeLux", "CS Italy",
//      "CS MEA North", "CS MEA South", "CS Southern Europe"
// To give access to multiple regions, use an array.
// To give full access to a specific user, set regions to []

window.ACCESS_CONFIG = {
  // Example entries — replace with real email addresses:
  // "firstname.lastname@sap.com": { regions: ["CS UKI"] },
  // "another.user@sap.com":       { regions: ["CS Nordic", "CS BeNeLux"] },
  // "admin.user@sap.com":         { regions: [] },  // full access
};
