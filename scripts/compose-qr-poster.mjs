import { readFile, writeFile } from "node:fs/promises";

const background = (await readFile("public/ul-lab-equipment-qr-poster-background.png")).toString("base64");
const qr = (await readFile("public/ul-lab-equipment-qr.svg")).toString("base64");

const poster = `<svg xmlns="http://www.w3.org/2000/svg" width="1254" height="1254" viewBox="0 0 1254 1254"><image href="data:image/png;base64,${background}" width="1254" height="1254"/><image href="data:image/svg+xml;base64,${qr}" x="344" y="402" width="566" height="566" preserveAspectRatio="xMidYMid meet"/></svg>`;
await writeFile("public/ul-lab-equipment-qr-poster.svg", poster);
