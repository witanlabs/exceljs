const {expect} = require('chai');
const {
  encodeStyle,
  encodeFont,
  encodeFill,
  encodeBorder,
  encodeAlignment,
  encodeProtection,
  encodeColor,
} = require('../../../lib/utils/style-fast-serialize');

describe('Style Fast Serialize', () => {
  describe('encodeColor', () => {
    it('should encode argb color', () => {
      expect(encodeColor({argb: 'FFFF0000'})).to.equal('a:FFFF0000');
    });

    it('should encode theme color', () => {
      expect(encodeColor({theme: 1})).to.equal('t:1');
      expect(encodeColor({theme: 2, tint: 0.5})).to.equal('t:2:0.5');
    });

    it('should encode indexed color', () => {
      expect(encodeColor({indexed: 64})).to.equal('i:64');
    });

    it('should return empty string for null/undefined', () => {
      expect(encodeColor(null)).to.equal('');
      expect(encodeColor(undefined)).to.equal('');
    });
  });

  describe('encodeFont', () => {
    it('should encode font with all properties', () => {
      const font = {
        name: 'Arial',
        size: 12,
        bold: true,
        italic: true,
        underline: true,
        strike: true,
        color: {argb: 'FF000000'},
      };
      const encoded = encodeFont(font);
      expect(encoded).to.include('n:Arial');
      expect(encoded).to.include('s:12');
      expect(encoded).to.include('b');
      expect(encoded).to.include('i');
      expect(encoded).to.include('u:single');
      expect(encoded).to.include('st');
      expect(encoded).to.include('c:a:FF000000');
    });

    it('should return empty string for null', () => {
      expect(encodeFont(null)).to.equal('');
    });
  });

  describe('encodeFill', () => {
    it('should encode pattern fill', () => {
      const fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {argb: 'FFFF0000'},
      };
      const encoded = encodeFill(fill);
      expect(encoded).to.include('t:pattern');
      expect(encoded).to.include('p:solid');
      expect(encoded).to.include('fg:a:FFFF0000');
    });

    it('should return empty string for null', () => {
      expect(encodeFill(null)).to.equal('');
    });
  });

  describe('encodeBorder', () => {
    it('should encode border with all sides', () => {
      const border = {
        top: {style: 'thin', color: {argb: 'FF000000'}},
        left: {style: 'thin'},
        bottom: {style: 'medium'},
        right: {style: 'thick'},
      };
      const encoded = encodeBorder(border);
      expect(encoded).to.include('t:s:thin');
      expect(encoded).to.include('l:s:thin');
      expect(encoded).to.include('b:s:medium');
      expect(encoded).to.include('r:s:thick');
    });

    it('should return empty string for null', () => {
      expect(encodeBorder(null)).to.equal('');
    });
  });

  describe('encodeAlignment', () => {
    it('should encode alignment', () => {
      const alignment = {
        horizontal: 'center',
        vertical: 'middle',
        wrapText: true,
        indent: 2,
      };
      const encoded = encodeAlignment(alignment);
      expect(encoded).to.include('h:center');
      expect(encoded).to.include('v:middle');
      expect(encoded).to.include('w');
      expect(encoded).to.include('i:2');
    });

    it('should return empty string for null', () => {
      expect(encodeAlignment(null)).to.equal('');
    });
  });

  describe('encodeProtection', () => {
    it('should encode protection', () => {
      expect(encodeProtection({locked: true})).to.equal('l:1');
      expect(encodeProtection({locked: false, hidden: true})).to.equal(
        'l:0|h:1'
      );
    });

    it('should return empty string for null', () => {
      expect(encodeProtection(null)).to.equal('');
    });
  });

  describe('encodeStyle', () => {
    it('should encode complete style model', () => {
      const model = {
        numFmt: '#,##0.00',
        font: {name: 'Arial', size: 10},
        fill: {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFF0000'}},
        border: {top: {style: 'thin'}},
        alignment: {horizontal: 'center'},
        protection: {locked: true},
      };
      const encoded = encodeStyle(model);
      expect(encoded).to.include('nf:#,##0.00');
      expect(encoded).to.include('fo:');
      expect(encoded).to.include('fi:');
      expect(encoded).to.include('bo:');
      expect(encoded).to.include('al:');
      expect(encoded).to.include('pr:');
    });

    it('should produce identical output for identical styles', () => {
      const style1 = {font: {name: 'Arial', size: 10, bold: true}};
      const style2 = {font: {name: 'Arial', size: 10, bold: true}};
      expect(encodeStyle(style1)).to.equal(encodeStyle(style2));
    });

    it('should produce different output for different styles', () => {
      const style1 = {font: {name: 'Arial', size: 10}};
      const style2 = {font: {name: 'Arial', size: 12}};
      expect(encodeStyle(style1)).to.not.equal(encodeStyle(style2));
    });

    it('should return empty string for null', () => {
      expect(encodeStyle(null)).to.equal('');
    });
  });
});
