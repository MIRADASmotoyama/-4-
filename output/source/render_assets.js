// scenes.js の SVG を PNG に書き出す
const sharp = require("sharp");
const fs = require("fs");
const S = require("./scenes");
const list = {
  city_bridge: S.cityBridge(), steel_mill: S.steelMill(), chip: S.chip(), port: S.port(),
  coil: S.coil(), hydrogen: S.hydrogen(), consumer: S.consumer(),
  blast_furnace: S.blastFurnace(), electric_furnace: S.electricFurnace(), discussion: S.discussion(),
  m_steel: S.metalSphere("#8E9BA6", "#F2F5F7", "#4B5864"),
  m_alu: S.metalSphere("#C9D1D8", "#FFFFFF", "#8A96A1"),
  m_copper: S.metalSphere("#C46A3A", "#F7C8A5", "#7A3518"),
  m_nickel: S.metalSphere("#A3A68F", "#F1F2E6", "#5F624F"),
  m_titan: S.metalSphere("#7E8DA6", "#E3E9F3", "#3E4A60"),
};
(async () => {
  fs.mkdirSync(__dirname + "/assets", { recursive: true });
  for (const [k, svg] of Object.entries(list)) {
    await sharp(Buffer.from(svg)).png().toFile(`${__dirname}/assets/${k}.png`);
  }
  console.log("assets:", Object.keys(list).length);
})();
