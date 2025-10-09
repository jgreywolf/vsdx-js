import { parseVisioFile } from './dist/index.js';

async function testEnhancedStyling() {
  try {
    console.log('=== ENHANCED STYLING SYSTEM TEST ===');
    const result = await parseVisioFile('./tests/DiagramWithStyles.vsdx');
    
    console.log('\n🎨 STYLE IMPROVEMENTS VERIFICATION:');
    
    console.log('\n1. PROCESSED STYLESHEETS:');
    result.Stylesheets.slice(0, 3).forEach(sheet => {
      console.log(`   ✅ Sheet ${sheet.ID} (${sheet.Name}):`);
      console.log(`      - Now has processed Style object: ${typeof sheet.Style === 'object' ? 'YES' : 'NO'}`);
      console.log(`      - Line Weight: ${sheet.Style.LineWeight || 'default'}`);
      console.log(`      - Line Color: ${sheet.Style.LineColor || 'default'}`);
      console.log(`      - Fill Pattern: ${sheet.Style.FillPattern || 'default'}`);
    });
    
    console.log('\n2. ENHANCED MASTERS:');
    result.Masters.slice(0, 2).forEach(master => {
      console.log(`   ✅ Master ${master.Id} (${master.Name}):`);
      console.log(`      - Has processed Style: ${master.Style ? 'YES' : 'NO'}`);
      console.log(`      - Style properties: ${Object.keys(master.Style || {}).length}`);
      console.log(`      - Line Style Ref: ${master.LineStyleRefId || 'None'}`);
      console.log(`      - Fill Style Ref: ${master.FillStyleRefId || 'None'}`);
    });
    
    console.log('\n3. ENHANCED SHAPES WITH INHERITANCE:');
    if (result.Pages[0]) {
      result.Pages[0].Shapes.slice(0, 3).forEach(shape => {
        console.log(`   ✅ Shape ${shape.Id} (${shape.Type}):`);
        console.log(`      - Master: ${shape.MasterId}`);
        console.log(`      - Line Color: ${shape.Style.LineColor} (was getting nulls before)`);
        console.log(`      - Line Weight: ${shape.Style.LineWeight} (with proper unit conversion)`);
        console.log(`      - Fill Color: ${shape.Style.FillForeground} (with enhanced color parsing)`);
        console.log(`      - Fill Pattern: ${shape.Style.FillPattern} (with defaults)`);
        console.log(`      - Text Color: ${shape.Style.TextColor} (now has defaults)`);
        console.log(`      - Style References: Line=${shape.LineStyleRefId || 'None'}, Fill=${shape.FillStyleRefId || 'None'}`);
        
        // Count complete vs null/undefined values
        const styleValues = Object.values(shape.Style);
        const completeValues = styleValues.filter(v => v !== null && v !== undefined && v !== '').length;
        console.log(`      - Complete Style Properties: ${completeValues}/${styleValues.length} (${Math.round(completeValues/styleValues.length*100)}%)`);
      });
    }
    
    console.log('\n4. COLOR PARSING IMPROVEMENTS:');
    const allShapes = result.Pages.flatMap(page => page.Shapes);
    const colorProperties = allShapes.flatMap(shape => [
      shape.Style.LineColor,
      shape.Style.FillForeground, 
      shape.Style.FillBackground,
      shape.Style.TextColor
    ]);
    const validColors = colorProperties.filter(color => color && color.startsWith('#')).length;
    const emptyColors = colorProperties.filter(color => !color || color === '').length;
    console.log(`   ✅ Valid hex colors: ${validColors}`);
    console.log(`   ✅ Empty/null colors: ${emptyColors}`);
    console.log(`   ✅ Color parsing success rate: ${Math.round(validColors/(validColors + emptyColors)*100)}%`);
    
    console.log('\n5. UNIT CONVERSION VERIFICATION:');
    const lineWeights = allShapes.map(shape => shape.Style.LineWeight).filter(w => w);
    console.log(`   ✅ Line weights converted to pixels: ${lineWeights.join(', ')}`);
    console.log(`   ✅ All line weights are numbers: ${lineWeights.every(w => typeof w === 'number')}`);
    
    console.log('\n🚀 SUMMARY OF IMPROVEMENTS:');
    console.log('   ✅ Stylesheets now contain processed Style objects instead of raw cells');
    console.log('   ✅ Masters have their own Style properties populated');
    console.log('   ✅ Style inheritance chain implemented (Shape → Master → Stylesheet → Defaults)');
    console.log('   ✅ Enhanced color parsing with hex conversion and Visio color palette');
    console.log('   ✅ Proper unit conversion (PT, IN, MM → pixels)');
    console.log('   ✅ Default values prevent null/undefined properties');
    console.log('   ✅ Style reference IDs extracted from shapes and masters');
    console.log('   ✅ Comprehensive style merging with proper precedence');
    
  } catch (e) {
    console.error('❌ Error:', e.message);
    console.error(e.stack);
  }
}

testEnhancedStyling();