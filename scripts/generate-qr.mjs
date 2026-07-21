import QRCode from "qrcode";
import { readFile, writeFile } from "node:fs/promises";

const siteUrl = "https://ul-lab-equipment-app.vercel.app";
const qr = await QRCode.toString(siteUrl, {
  type: "svg",
  width: 1200,
  margin: 5,
  errorCorrectionLevel: "H",
  color: { dark: "#171717", light: "#FFFFFF" },
});
const logo = (await readFile("public/assets/ul-solutions-logo.png")).toString("base64");
const logoOverlay = `<g><rect x="16" y="16" width="15" height="15" rx="1" fill="#fff"/><image href="data:image/png;base64,${logo}" x="17" y="21" width="13" height="5" preserveAspectRatio="xMidYMid meet"/></g>`;
await writeFile("public/ul-lab-equipment-qr.svg", qr.replace("</svg>", `${logoOverlay}</svg>`));
