# Verbatim AI HTML Report Documentation

## Overview
The Verbatim AI HTML report is designed to provide a clear, professional comparison between draft content and live website content. This document outlines the structure, styling decisions, and guidelines for interpreting the reports.

## HTML Structure

### Container Hierarchy
```html
<div class='report-container'>
    <div class='header'>
        <!-- Title and color key -->
    </div>
    <div class='page-info'>
        <!-- Document metadata -->
    </div>
    <div class='content-comparison'>
        <div class='column-headers'>
            <!-- Column labels -->
        </div>
        <div class='content-container'>
            <!-- Content blocks -->
        </div>
    </div>
</div>
```

### Key Components

1. **Header Section**
   - Contains the main title and color key
   - Uses a clean, professional font stack (Roboto/Arial/sans-serif)
   - Color key explains the meaning of different block colors

2. **Page Info Section**
   - Displays document metadata (filename, URL, title, meta description)
   - Shows similarity score and visual indicator
   - Groups related information in a visually distinct section

3. **Content Comparison**
   - Two-column layout with clear headers
   - Content blocks maintain consistent spacing and alignment
   - Color-coded blocks for different content states

## CSS Guidelines

### Core Design Principles

1. **Typography**
   - Primary font: Roboto
   - Fallback fonts: Arial, sans-serif
   - Base line height: 1.6
   - Headers use color #2c3e50 for better readability

2. **Color Scheme**
   ```css
   /* Background colors */
   .matched-content  { background-color: #e8f5e9; }  /* Light green */
   .missing-content  { background-color: #ffebee; }  /* Light red */
   .current-content { background-color: #e3f2fd; }  /* Light blue */
   
   /* Text colors */
   .matched-text  { color: #28a745; }  /* Green */
   .missing-text  { color: #dc3545; }  /* Red */
   .current-text  { color: #007bff; }  /* Blue */
   ```

3. **Layout**
   - Max width: 1400px
   - Responsive padding: 20px body, 30px container
   - Consistent spacing using 15px gaps
   - Flexbox for alignment and distribution

4. **Visual Hierarchy**
   - Subtle shadows (0 2px 4px rgba(0,0,0,0.1))
   - Rounded corners (5px border-radius)
   - Light borders for definition
   - Background contrast for sections

### Important CSS Classes

```css
.report-container {
    /* Main container */
    max-width: 1400px;
    margin: 0 auto;
    background-color: white;
    padding: 30px;
    border-radius: 8px;
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
}

.content-block {
    /* Individual content blocks */
    flex: 1;
    padding: 15px;
    border-radius: 5px;
    white-space: pre-wrap;
    word-break: break-word;
    min-height: 50px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
}
```

## Modification Guidelines

### Do's
1. Maintain the existing color scheme for consistency
2. Keep the two-column layout for content comparison
3. Preserve whitespace handling (pre-wrap) in content blocks
4. Use the established CSS class hierarchy
5. Keep the visual hierarchy with headers and sections

### Don'ts
1. Don't remove the color key or similarity indicators
2. Don't change the font stack without fallbacks
3. Don't modify the content block structure without testing text wrapping
4. Don't remove the container max-width (important for readability)
5. Don't change the spacing system without updating all related elements

## Content Block States

1. **Matched Content**
   - Green background (#e8f5e9)
   - Green border (#c8e6c9)
   - Used for content that matches between draft and live

2. **Missing Content**
   - Red background (#ffebee)
   - Red border (#ffcdd2)
   - Used for content in draft but missing from live

3. **Current Content**
   - Blue background (#e3f2fd)
   - Blue border (#bbdefb)
   - Used for content on live site but not in draft

4. **Placeholder**
   - Light gray background (#f8f9fa)
   - Dashed border (#dee2e6)
   - Used for empty states with descriptive messages

## Future Enhancements

When making changes to the report format, consider:

1. **Accessibility**
   - Maintain color contrast ratios
   - Keep text sizes readable
   - Ensure keyboard navigation works

2. **Responsiveness**
   - Test on different screen sizes
   - Consider adding breakpoints for mobile
   - Maintain readability at all sizes

3. **Performance**
   - Keep CSS selectors simple
   - Minimize nested elements
   - Use efficient CSS properties

## Testing Changes

Before implementing changes:

1. Test with various content lengths
2. Verify text wrapping behavior
3. Check spacing consistency
4. Validate color contrast
5. Test in multiple browsers

## Example Content Block

```html
<div class='content-row'>
    <div class='content-block matched-content'>
        <!-- Draft content -->
    </div>
    <div class='content-block matched-content'>
        <!-- Live content -->
    </div>
</div>
```

This structure ensures content blocks maintain proper alignment and spacing while preserving the visual hierarchy of the report. 