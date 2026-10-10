// Rewrites the host in manifest.xml (whatever it currently is) to a new HTTPS host.
// Usage: npm run set-host -- https://YOUR-HOST/tms-tracker-addin
import { readFileSync, writeFileSync } from "node:fs";
const host = (process.argv[2] || "").replace(/\/+$/, "");
if (!/^https:\/\/[^\s]+$/.test(host)) { console.error("Give the HTTPS host, e.g. npm run set-host -- https://intranet.example.org/tms-tracker-addin"); process.exit(1); }
const origin = new URL(host).origin;
let xml = readFileSync("manifest.xml", "utf8");
const m = xml.match(/<SourceLocation DefaultValue="(https:\/\/[^"]+)\/taskpane\.html"/);
if (!m) { console.error("Could not find the current host in manifest.xml"); process.exit(1); }
const oldHost = m[1], oldOrigin = new URL(oldHost).origin;
xml = xml.split(`<AppDomain>${oldOrigin}</AppDomain>`).join(`<AppDomain>${origin}</AppDomain>`).split(oldHost).join(host);
writeFileSync("manifest.xml", xml);
console.log(`manifest.xml now points to ${host} (was ${oldHost})`);
