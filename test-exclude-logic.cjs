// Test the exclude logic for account 1646
const testExcludes = () => {
  // Simulating the exclude logic from SelectionFirstClassifier.buildResolver
  
  const rules = {
    ranges: [{ from: 1610, to: 1659 }],
    includes: [1600],
    excludes: [1646, 1647, 1646.1, 1647.1]
  };
  
  const includes = new Set(rules.includes.map(String));
  const excludes = new Set(rules.excludes.map(String));
  const ranges = rules.ranges || [];
  
  console.log("\n=== Testing PPE Cost Exclude Logic ===\n");
  console.log("Rules:", JSON.stringify(rules, null, 2));
  console.log("\nExcludes Set:", Array.from(excludes));
  
  const testAccounts = ['1600', '1610', '1646', '1647', '1646.1', '1647.1', '1650'];
  
  testAccounts.forEach(codeStr => {
    const codeNum = Number.parseFloat(codeStr || '0');
    
    const byInclude = includes.size > 0 && includes.has(codeStr);
    const byRange = ranges.length > 0 && Number.isFinite(codeNum) && 
                   ranges.some(r => codeNum >= r.from && codeNum <= r.to);
    const excluded = excludes.has(codeStr);
    
    const matched = (byInclude || byRange) && !excluded;
    
    console.log(`\nAccount ${codeStr}:`);
    console.log(`  codeNum: ${codeNum}`);
    console.log(`  byInclude: ${byInclude}`);
    console.log(`  byRange: ${byRange}`);
    console.log(`  excluded: ${excluded}`);
    console.log(`  => MATCHED: ${matched}`);
  });
};

testExcludes();
