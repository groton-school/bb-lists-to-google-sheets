const config = require('@battis/webpack-typescript-gas');

c = config({ root: __dirname });

/*
c.mode = 'development';
c.devtool = 'inline-source-map';
*/

module.exports = c;
