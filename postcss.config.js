const tailwindcss = require("tailwindcss");
const autoprefixer = require("autoprefixer");
module.exports = {
  plugins: {
    'tailwindcss/nesting': 'postcss-nesting',
    tailwindcss,
    autoprefixer,
    'postcss-preset-env': {
      features: { 'nesting-rules': false },
    },
  }
}