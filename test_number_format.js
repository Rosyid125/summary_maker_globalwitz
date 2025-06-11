// Test file untuk verifikasi format angka
const { parseNumber } = require('./src/utils');

console.log('Testing European format (koma sebagai desimal):');
console.log('1.234,56 =>', parseNumber('1.234,56', 'EUROPEAN')); // Expected: 1234.56
console.log('123,45 =>', parseNumber('123,45', 'EUROPEAN'));     // Expected: 123.45
console.log('1234 =>', parseNumber('1234', 'EUROPEAN'));         // Expected: 1234

console.log('\nTesting American format (titik sebagai desimal):');
console.log('1,234.56 =>', parseNumber('1,234.56', 'AMERICAN')); // Expected: 1234.56
console.log('123.45 =>', parseNumber('123.45', 'AMERICAN'));     // Expected: 123.45
console.log('1234 =>', parseNumber('1234', 'AMERICAN'));         // Expected: 1234

console.log('\nTesting edge cases:');
console.log('Empty string =>', parseNumber('', 'EUROPEAN'));     // Expected: 0
console.log('Invalid =>', parseNumber('abc', 'EUROPEAN'));       // Expected: 0
console.log('Number =>', parseNumber(123.45, 'EUROPEAN'));       // Expected: 123.45
