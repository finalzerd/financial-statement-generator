// Diagnostic script to check intangible assets calculation
console.log('Testing intangible assets calculation:');
console.log('');

// Simulate the data
const costCurrent = 12000;
const accumAmortCurrent = 2827.4;

console.log('Cost (current):', costCurrent);
console.log('Accumulated Amortization (current):', accumAmortCurrent);
console.log('');

// Test the calculation
const netBookValue = costCurrent - accumAmortCurrent;
console.log('Net Book Value (Cost - Accum):', netBookValue);
console.log('Expected: 9172.6');
console.log('');

// What you're seeing
const wrongValue = costCurrent + accumAmortCurrent;
console.log('Wrong calculation (Cost + Accum):', wrongValue);
console.log('This matches what you see: 14827.4');
console.log('');

// Check if it's a sign issue
console.log('If accumulated amortization was negative in TB:');
const accumNegative = -2827.4;
const withNegative = costCurrent + accumNegative;
console.log(`  Cost + (${accumNegative}) = ${withNegative}`);
console.log('  This would give correct result without needing Math.abs()');
