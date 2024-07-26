const fs = require('fs');

function readConfig(configPath) {
  const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
  return {
    username: config.username,
    password: config.password,
    excelPath: config.excelPath
  };
}

module.exports = { readConfig };
