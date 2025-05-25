# PowerPoint Layout Designer with AI Content Generation

A comprehensive web-based tool that combines visual PowerPoint slide layout design with AI-powered content generation. Design custom layouts with templates, generate presentation content with OpenAI, and create professional PowerPoint files with precise formatting and bullet points.

## Overview

This application provides a complete presentation creation workflow:
- **Visual Layout Design**: Drag-and-drop interface with multiple templates
- **AI Content Generation**: OpenAI integration for slide planning and content creation
- **Template Selection**: Standard and custom layouts with image placeholders
- **Interactive Content Planning**: Review and edit AI-generated slide outlines
- **Professional Output**: Creates PPTX files with proper bullet formatting and image placeholders

## Features

### Template Selection
- **Standard Content Template**: Traditional single-column layout
- **My Content Template (with Image)**: Two-column layout with image placeholder on the right
- **Template Switching**: Easy switching between templates with preserved customizations
- **Smart Layouts**: Templates automatically adjust for title and content slides

### Visual Layout Editor
- **Drag & Drop Interface**: Move elements around the slide canvas
- **Resize Handles**: Precise sizing with visual feedback
- **Real-time Preview**: See exactly how your slide will look
- **Element Types**: Title boxes, text boxes, and image placeholders
- **Formatting Options**: Font selection, sizes, list types, and positioning
- **Template-Based Design**: Start with professional templates or create custom layouts

### AI Content Generation
- **Topic-Based Generation**: Enter any topic to generate presentation outlines
- **Draft Planning**: AI creates logical slide structure with titles
- **Interactive Editing**: Modify, add, delete, and reorder slides
- **Full Content Generation**: AI creates detailed bullet-point content for each slide
- **Content Refinement**: Edit generated content before final presentation creation

### Professional Output
- **Proper Bullet Points**: Generates actual bullet formatting in PowerPoint
- **16:9 Aspect Ratio**: Professional widescreen format
- **Font Scaling**: Maintains exact font sizes (18pt displays as 18pt)
- **Image Placeholders**: Creates styled rectangles with insertion instructions
- **Layout Preservation**: Your custom layouts are applied to all slides

### Code & File Generation
- **Live PPTX Creation**: Generate actual PowerPoint files with your content and layouts
- **Python Code Export**: Get clean python-pptx scripts for automation
- **Layout Management**: Save/load custom layouts for reuse
- **Template Integration**: Generated code includes template configurations

## Installation

### Prerequisites
- Python 3.x
- Flask web framework
- OpenAI API key
- python-pptx library

### Setup

1. **Install dependencies:**
   ```bash
   pip install flask openai python-pptx
   ```

2. **Set up OpenAI API key:**
   ```bash
   export OPENAI_API_KEY="your-api-key-here"
   ```

3. **Run the application:**
   ```bash
   python3 app.py
   ```

4. **Open your browser:**
   Navigate to `http://localhost:5000`

## Usage Workflow

### 1. Select Your Template
- Choose between "Standard Content Template" or "My Content Template (with Image)"
- Template selection affects the layout of content slides
- Switch templates anytime - your customizations are preserved

### 2. Design Your Layouts
- Use the visual editor to customize your selected template
- Add title boxes, text boxes, and image placeholders
- Position and size elements precisely
- Configure fonts, sizes, and formatting options
- Toggle between Title Slide and Content Slide types

### 3. Generate Content with AI
- Enter your presentation topic in the sidebar
- Click "Generate Draft" to create an AI-powered slide outline
- Review the generated slide titles and structure

### 4. Refine Your Content
- Click "Content Planner" to open the interactive editor
- Edit slide titles and add/remove slides
- Reorder slides by moving them up or down
- Click "Generate Content" to create detailed slide content
- Edit the AI-generated content as needed

### 5. Create Your Presentation
- Once satisfied with content and layouts, click "Create PowerPoint"
- Download the complete PPTX file with your custom layouts and AI content
- The presentation includes proper bullet points and image placeholders

## Templates

### Standard Content Template
- Traditional single-column layout
- Full-width title at top
- Full-width content area below
- Perfect for text-heavy presentations

### My Content Template (with Image)
- Two-column layout design
- Full-width title across the top (40pt font)
- Text content on the left (28pt font with bullets)
- Image placeholder on the right side
- Matches the dimensions you provided in your Python code

## Technical Details

### Coordinate System
- Canvas: 800px × 450px (16:9 aspect ratio)
- PowerPoint slide: 13.333" × 7.5" (16:9 widescreen)
- Precise inch-based positioning

### Font Handling
- Disabled auto-sizing to preserve exact font sizes
- Force run creation for proper font application
- Manual bullet character addition for reliable formatting

### Image Handling
- Creates styled rectangle placeholders with gray borders
- Displays "[INSERT IMAGE HERE]" text
- Generated code includes try/catch for actual image loading

### AI Integration
- Uses OpenAI GPT-3.5-turbo for content generation
- Structured prompts for consistent bullet-point format
- Each line becomes a separate bullet point automatically

## API Endpoints

### Layout Designer
- `GET /` - Main application interface
- `POST /generate_code` - Generate python-pptx code from layout data
- `POST /download_code` - Download generated code as file

### AI Content Generation
- `POST /generate_draft` - Generate slide outline from topic
- `POST /generate_content` - Generate detailed content for slide outline
- `POST /create_presentation` - Create final PPTX with layouts and content

## Example Workflow

### 1. Select Template and Create Layout
```
Choose "My Content Template (with Image)"
Customize positioning if needed
Configure fonts and formatting
```

### 2. Generate Content
```
Topic: "Introduction to Machine Learning"
AI generates: Title slide + 5 content slides with topics
Edit/refine the slide outline as needed
Generate detailed bullet-point content for each slide
```

### 3. Download Professional Presentation
```
System combines your template with AI content
Downloads PPTX with proper bullets and image placeholders
Ready for presentation or further editing in PowerPoint
```

## File Structure
```
powerpoint-layout-designer/
├── app.py                 # Flask backend with template support
├── templates/
│   └── index.html        # Web interface with template selection
├── static/               # Static assets (if needed)
├── gitignore/           # Large files (excluded from git)
└── README.md            # This file
```

## Dependencies
- **Backend**: Flask, OpenAI Python SDK, python-pptx
- **Frontend**: Vanilla JavaScript, HTML5, CSS3
- **AI**: OpenAI GPT-3.5-turbo API

## Environment Variables
- `OPENAI_API_KEY` - Your OpenAI API key (required for content generation)

## Troubleshooting

### Common Issues

**Bullet points not appearing:**
- The app now manually adds bullet characters (•) for reliable formatting
- Each line of AI-generated content becomes a separate bullet point

**Font sizes appear small:**
- Auto-sizing has been disabled to preserve exact font sizes
- 18pt fonts now display as actual 18pt in PowerPoint

**Template not loading:**
- Check browser console for JavaScript errors
- Clear browser cache and reload
- Ensure template selector shows the correct option

**OpenAI API errors:**
- Ensure `OPENAI_API_KEY` environment variable is set
- Check API key validity and account credits
- Verify internet connection for API calls

**Image placeholders not showing:**
- Images are now created as styled rectangles with borders
- Look for gray rectangles with "[INSERT IMAGE HERE]" text
- Replace with actual images in PowerPoint after download

## Contributing

This project combines visual design tools with AI content generation and professional template support for automated presentation creation. The template system allows for flexible layout options while maintaining professional formatting standards.

## License

Personal use project. Modify as needed for your workflows.

## Related Technologies

- **python-pptx**: PowerPoint file generation with precise formatting
- **OpenAI API**: AI content generation with structured prompts
- **Flask**: Web framework for the layout designer
- **HTML5 Canvas**: Visual layout editor with template support