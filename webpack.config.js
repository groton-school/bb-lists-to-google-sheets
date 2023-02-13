const path = require('path');
const config = require('@battis/webpack-typescript-gas');

config.output.path = path.join(__dirname, 'build');

module.exports = config;
