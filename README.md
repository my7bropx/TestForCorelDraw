# CorelDRAW Fill Curve Tool

A powerful and user-friendly CorelDRAW macro tool that fills open or closed curves with selected elements while preserving all original properties.

## Features

- **Flexible Curve Support**: Works with both open and closed curves
- **Preserve Element Properties**: Maintains original size, angle, colors, and all attributes
- **Adjustable Spacing**: Control horizontal and vertical spacing between elements
- **Multiple Fill Patterns**:
  - Grid Pattern (regular rows and columns)
  - Hexagonal Pattern (honeycomb style)
- **Multiple Elements**: Can cycle through multiple fill elements for varied patterns
- **User-Friendly Interface**: Easy-to-use dialogs and prompts
- **Wide Compatibility**: Works with CorelDRAW 2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025, and 2026

## Version Information

Two versions are provided:

1. **FillCurveTool.gms** - Advanced version with dialog interface
2. **FillCurveTool_Simple.gms** - Simplified version using InputBox (maximum compatibility)

Both versions offer the same core functionality. Use the Simple version if you experience any compatibility issues.

## Installation

### Method 1: Quick Install (Recommended)

1. Download the `.gms` file to your computer
2. Open CorelDRAW
3. Go to `Tools` → `Visual Basic` → `Play` (or press `Alt+F11` then `F5`)
4. Click `Browse` and select the downloaded `.gms` file
5. Click `Run`

### Method 2: Install as Macro

1. Open CorelDRAW
2. Press `Alt+F11` to open the Visual Basic Editor
3. In the Project Explorer, find `GlobalMacros` or your document name
4. Right-click on the project → `Insert` → `Module`
5. Copy and paste the macro code into the module
6. Save and close the VB Editor

### Method 3: Add to Toolbar (For Frequent Use)

1. Install the macro using Method 2
2. Go to `Tools` → `Customization` → `Commands`
3. Select `Macros` from the category list
4. Find your macro (e.g., `FillCurveSimple` or `FillCurveWithElements`)
5. Drag it to your toolbar
6. Click `OK`

## How to Use

### Basic Usage

1. **Create your elements**:
   - Create or select the element(s) you want to use as fill (e.g., circles, stars, logos)
   - These can be any CorelDRAW objects (shapes, text, groups, etc.)

2. **Create your container**:
   - Create or select the curve you want to fill (can be open or closed)
   - The container should be larger than the fill elements

3. **Select objects**:
   - Select ALL objects: fill element(s) + container curve
   - The largest object will automatically be used as the container

4. **Run the macro**:
   - `FillCurveSimple` - For the simple version with prompts
   - `QuickFill` - For instant fill with default settings (0% spacing, grid pattern)
   - `FillCurveWithElements` - For the advanced version with dialog

5. **Configure settings** (if using FillCurveSimple):
   - Choose pattern type (1 = Grid, 2 = Hexagonal)
   - Set horizontal spacing percentage
   - Set vertical spacing percentage

6. **Done!**: The macro will fill your curve with the selected elements

### Understanding Spacing

The spacing values control the gap between elements:

- **0%**: Elements touch each other (no gap)
- **50%**: Half element width/height gap
- **100%**: Full element width/height gap
- **Negative values**: Elements overlap (e.g., -50% = 50% overlap)

**Examples**:
- Touching circles: `H: 0%, V: 0%`
- Small gaps: `H: 10%, V: 10%`
- Spaced out: `H: 50%, V: 50%`
- Overlapping: `H: -20%, V: -20%`

## Pattern Types

### Grid Pattern
Regular rows and columns, like a checkerboard. Best for:
- Simple, uniform fills
- Geometric patterns
- Organized layouts

### Hexagonal Pattern
Honeycomb-style pattern with offset rows. Best for:
- Natural, organic looks
- Dense packing
- Artistic effects

## Examples and Use Cases

### Example 1: Fill Circle with Stars
```
1. Create a star (small)
2. Create a circle (large)
3. Select both
4. Run FillCurveSimple
5. Choose Grid pattern, 10% spacing
Result: Circle filled with evenly spaced stars
```

### Example 2: Hexagonal Logo Pattern
```
1. Create/import your logo (small)
2. Create an outline shape
3. Select both
4. Run FillCurveSimple
5. Choose Hexagonal pattern, 0% spacing
Result: Dense honeycomb fill of logos
```

### Example 3: Multiple Element Fill
```
1. Create circle (red, small)
2. Create square (blue, small)
3. Create triangle (green, small)
4. Create container curve (large)
5. Select all four objects
6. Run FillCurveSimple
7. Choose Grid pattern, 20% spacing
Result: Alternating pattern of circles, squares, and triangles
```

### Example 4: Text Pattern Fill
```
1. Create text "SALE" (small)
2. Rotate it 45 degrees
3. Create rectangle (large)
4. Select both
5. Run QuickFill (for quick result)
Result: Rectangle filled with rotated "SALE" text
```

## Tips and Best Practices

1. **Element Size**: Keep fill elements much smaller than the container for best results
2. **Performance**: For very large fills (thousands of elements), be patient - the macro optimizes automatically
3. **Overlapping**: Use negative spacing values to create interesting overlap effects
4. **Multiple Elements**: Select multiple small elements to create varied, interesting patterns
5. **Rotation**: Rotate your fill elements before running the macro to create angled patterns
6. **Groups**: You can use grouped objects as fill elements
7. **Colors**: All colors, gradients, and effects are preserved
8. **Undo**: The entire fill operation is one undo step - if you don't like it, just press Ctrl+Z

## Troubleshooting

### "Select exactly TWO shapes" error
- **Solution**: Make sure you've selected at least 2 objects (fill element + container)

### No elements appear
- **Solution**:
  - Check that your container is larger than fill elements
  - Verify the container is a valid curve
  - Try different spacing values

### Too many/few elements
- **Solution**: Adjust spacing percentage values
  - Increase spacing = fewer elements
  - Decrease spacing (or negative) = more elements

### Macro won't run
- **Solution**:
  - Make sure macros are enabled in CorelDRAW
  - Try the Simple version (FillCurveTool_Simple.gms)
  - Check that you're using CorelDRAW 2018 or later

### Elements outside the curve
- **Solution**: This shouldn't happen, but if it does:
  - Try using a closed curve instead of an open one
  - Verify the curve doesn't have broken segments

## Technical Notes

### Compatibility
- Tested on CorelDRAW 2018-2026
- Uses standard VBA commands for maximum compatibility
- No external dependencies required

### Performance
- Automatically optimizes for large fills
- Uses command grouping for single undo
- Typical fill of 100-500 elements takes 1-5 seconds

### Limitations
- Very complex curves may slow down the fill operation
- Maximum recommended: ~5000 elements per fill
- Works best with closed curves (open curves supported but may have edge cases)

## File Descriptions

- **FillCurveTool.gms** - Advanced version with dialog interface (if supported)
- **FillCurveTool_Simple.gms** - Maximum compatibility version with InputBox prompts
- **Fill with Specific Elements** - Original basic version (legacy)
- **tested** - Original hex pattern version (legacy)

## Quick Reference Card

### Available Macros

| Macro Name | Description |
|------------|-------------|
| `FillCurveSimple` | Simple version with prompts (recommended) |
| `QuickFill` | Instant fill with default settings (no prompts) |
| `FillCurveWithElements` | Advanced version with dialog |

### Keyboard Shortcuts

After adding to toolbar, you can assign keyboard shortcuts via:
`Tools` → `Customization` → `Commands` → Right-click macro → `Properties`

## Support and Feedback

For issues, suggestions, or questions:
- Check the troubleshooting section above
- Review the examples for common use cases
- Ensure you're using CorelDRAW 2018 or later

## Version History

### Version 2.0 (Current)
- Added support for open curves
- Multiple fill elements support
- Adjustable spacing controls
- Two pattern types (Grid, Hexagonal)
- Improved user interface
- Enhanced compatibility (2018-2026)
- QuickFill macro for instant results

### Version 1.0 (Legacy)
- Basic hexagonal fill
- Closed curves only
- Fixed spacing

## License

See LICENSE file for details.

## Credits

Created for CorelDRAW users who need flexible, powerful fill tools while maintaining complete control over their design elements.

---

**Enjoy creating amazing patterns with the Fill Curve Tool!**
