/**
 * Fast style serialization for efficient style caching.
 *
 * This module provides fast string-based serialization of style objects
 * for use as cache keys. Unlike JSON.stringify, it produces deterministic
 * output and is optimized for the specific structure of Excel style objects.
 *
 * Original PR: https://github.com/exceljs/exceljs/pull/2867
 */

const DELIMITER = '|';

/**
 * Encode a font object to a string key
 * @param {Object} font
 * @returns {string}
 */
function encodeFont(font) {
  if (!font) return '';
  const parts = [];
  if (font.name) parts.push(`n:${font.name}`);
  if (font.size) parts.push(`s:${font.size}`);
  if (font.family) parts.push(`f:${font.family}`);
  if (font.scheme) parts.push(`sc:${font.scheme}`);
  if (font.charset) parts.push(`ch:${font.charset}`);
  if (font.color) parts.push(`c:${encodeColor(font.color)}`);
  if (font.bold) parts.push('b');
  if (font.italic) parts.push('i');
  if (font.underline) parts.push(`u:${font.underline === true ? 'single' : font.underline}`);
  if (font.vertAlign) parts.push(`va:${font.vertAlign}`);
  if (font.strike) parts.push('st');
  if (font.outline) parts.push('o');
  return parts.join(DELIMITER);
}

/**
 * Encode a color object to a string key
 * @param {Object} color
 * @returns {string}
 */
function encodeColor(color) {
  if (!color) return '';
  if (color.argb) return `a:${color.argb}`;
  if (color.theme !== undefined) {
    let result = `t:${color.theme}`;
    if (color.tint !== undefined) result += `:${color.tint}`;
    return result;
  }
  if (color.indexed !== undefined) return `i:${color.indexed}`;
  return '';
}

/**
 * Encode a fill object to a string key
 * @param {Object} fill
 * @returns {string}
 */
function encodeFill(fill) {
  if (!fill) return '';
  const parts = [`t:${fill.type || 'pattern'}`];
  if (fill.pattern) parts.push(`p:${fill.pattern}`);
  if (fill.fgColor) parts.push(`fg:${encodeColor(fill.fgColor)}`);
  if (fill.bgColor) parts.push(`bg:${encodeColor(fill.bgColor)}`);
  // Gradient fill
  if (fill.gradient) {
    parts.push(`g:${fill.gradient}`);
    if (fill.degree !== undefined) parts.push(`d:${fill.degree}`);
    if (fill.center) parts.push(`cn:${fill.center.left},${fill.center.top}`);
    if (fill.stops) {
      const stops = fill.stops.map(s => `${s.position}:${encodeColor(s.color)}`).join(';');
      parts.push(`st:${stops}`);
    }
  }
  return parts.join(DELIMITER);
}

/**
 * Encode a border side to a string key
 * @param {Object} side
 * @returns {string}
 */
function encodeBorderSide(side) {
  if (!side) return '';
  const parts = [];
  if (side.style) parts.push(`s:${side.style}`);
  if (side.color) parts.push(`c:${encodeColor(side.color)}`);
  return parts.join(DELIMITER);
}

/**
 * Encode a border object to a string key
 * @param {Object} border
 * @returns {string}
 */
function encodeBorder(border) {
  if (!border) return '';
  const parts = [];
  if (border.top) parts.push(`t:${encodeBorderSide(border.top)}`);
  if (border.left) parts.push(`l:${encodeBorderSide(border.left)}`);
  if (border.bottom) parts.push(`b:${encodeBorderSide(border.bottom)}`);
  if (border.right) parts.push(`r:${encodeBorderSide(border.right)}`);
  if (border.diagonal) {
    parts.push(`d:${encodeBorderSide(border.diagonal)}`);
    if (border.diagonalUp) parts.push('du');
    if (border.diagonalDown) parts.push('dd');
  }
  return parts.join(DELIMITER);
}

/**
 * Encode an alignment object to a string key
 * @param {Object} alignment
 * @returns {string}
 */
function encodeAlignment(alignment) {
  if (!alignment) return '';
  const parts = [];
  if (alignment.horizontal) parts.push(`h:${alignment.horizontal}`);
  if (alignment.vertical) parts.push(`v:${alignment.vertical}`);
  if (alignment.wrapText) parts.push('w');
  if (alignment.shrinkToFit) parts.push('s');
  if (alignment.indent !== undefined) parts.push(`i:${alignment.indent}`);
  if (alignment.readingOrder !== undefined) parts.push(`ro:${alignment.readingOrder}`);
  if (alignment.textRotation !== undefined) parts.push(`tr:${alignment.textRotation}`);
  return parts.join(DELIMITER);
}

/**
 * Encode a protection object to a string key
 * @param {Object} protection
 * @returns {string}
 */
function encodeProtection(protection) {
  if (!protection) return '';
  const parts = [];
  if (protection.locked !== undefined) parts.push(`l:${protection.locked ? 1 : 0}`);
  if (protection.hidden !== undefined) parts.push(`h:${protection.hidden ? 1 : 0}`);
  return parts.join(DELIMITER);
}

/**
 * Encode a complete style model to a string key
 * @param {Object} model - The style model
 * @returns {string}
 */
function encodeStyle(model) {
  if (!model) return '';
  const parts = [];
  if (model.numFmt) parts.push(`nf:${model.numFmt}`);
  if (model.font) parts.push(`fo:${encodeFont(model.font)}`);
  if (model.fill) parts.push(`fi:${encodeFill(model.fill)}`);
  if (model.border) parts.push(`bo:${encodeBorder(model.border)}`);
  if (model.alignment) parts.push(`al:${encodeAlignment(model.alignment)}`);
  if (model.protection) parts.push(`pr:${encodeProtection(model.protection)}`);
  return parts.join('~');
}

module.exports = {
  encodeStyle,
  encodeFont,
  encodeFill,
  encodeBorder,
  encodeAlignment,
  encodeProtection,
  encodeColor,
};
