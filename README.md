# PowerPoint Layout Designer with AI Content Generation

A comprehensive web-based tool that combines visual PowerPoint slide layout design with AI-powered content generation. Design custom layouts visually, generate presentation content with OpenAI, and create professional PowerPoint files with precise formatting.

## Overview

This application provides a complete presentation creation workflow:
- **Visual Layout Design**: Drag-and-drop interface for custom slide layouts
- **AI Content Generation**: OpenAI integration for slide planning and content creation
- **Interactive Content Planning**: Review and edit AI-generated slide outlines
- **Seamless Integration**: Combines custom layouts with AI content into downloadable PPTX files

## Features

### Visual Layout Editor
- **Drag & Drop Interface**: Move elements around the slide canvas
- **Resize Handles**: Precise sizing with visual feedback
- **Real-time Preview**: See exactly how your slide will look
- **Element Types**: Title boxes, text boxes, and image placeholders
- **Formatting Options**: Font selection, sizes, list types, and positioning

### AI Content Generation
- **Topic-Based Generation**: Enter any topic to generate presentation outlines
- **Draft Planning**: AI creates logical slide structure with titles
- **Interactive Editing**: Modify, add, delete, and reorder slides
- **Full Content Generation**: AI creates detailed bullet-point content for each slide
- **Content Refinement**: Edit generated content before final presentation creation

### Slide Layout Types
- **Title Slide**: Special formatting for presentation headers
- **Content Slide**: Standard layout for presentation content
- **Custom Positioning**: Precise element placement in inches

### Code & File Generation
- **Live PPTX Creation**: Generate actual PowerPoint files with your content and layouts
- **Python Code Export**: Get clean python-pptx scripts for automation
- **Layout Management**: Save/load custom layouts for reuse

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

### 1. Design Your Layouts
- Use the visual editor to create custom slide layouts
- Add title boxes, text boxes, and image placeholders
- Position and size elements precisely
- Configure fonts, sizes, and formatting options
- Toggle between Title Slide and Content Slide types

### 2. Generate Content with AI
- Enter your presentation topic in the sidebar
- Click "Generate Draft" to create an AI-powered slide outline
- Review the generated slide titles and structure

### 3. Refine Your Content
- Click "Content Planner" to open the interactive editor
- Edit slide titles and add/remove slides
- Reorder slides by moving them up or down
- Click "Generate Content" to create detailed slide content
- Edit the AI-generated content as needed

### 4. Create Your Presentation
- Once satisfied with content and layouts, click "Create PowerPoint"
- Download the complete PPTX file with your custom layouts and AI content
- The presentation combines your visual design with the generated content

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

### 1. Create a Layout
```
Add title box at top (Arial 24pt)
Add content text box below (Calibri 18pt, bullet points)
Position elements precisely on 10" × 5.625" slide
```

### 2. Generate Content
```
Topic: "Introduction to Machine Learning"
AI generates: Title slide + 5 content slides with topics
Edit/refine the slide outline as needed
Generate detailed content for each slide
```

### 3. Download Presentation
```
System combines your custom layout with AI content
Downloads professional PPTX file ready for presentation
```

## Technical Details

### Coordinate System
- Canvas: 800px × 450px (16:9 aspect ratio)
- Conversion: 80 pixels = 1 inch
- PowerPoint slide: 10" × 5.625"

### AI Integration
- Uses OpenAI GPT-3.5-turbo for content generation
- Structured prompts for consistent output format
- Error handling for API failures
- Supports topic-based outline generation and detailed content creation

### File Structure
```
powerpoint-layout-designer/
├── app.py                 # Flask backend with AI integration
├── templates/
│   └── index.html        # Web interface with content planner
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

**OpenAI API errors:**
- Ensure `OPENAI_API_KEY` environment variable is set
- Check API key validity and account credits
- Verify internet connection for API calls

**Flask not found:**
```bash
pip install flask
```

**python-pptx not found:**
```bash
pip install python-pptx
```

**Generated presentation errors:**
- Ensure layouts have at least one text element
- Check that slide content is properly formatted
- Verify element positioning is within slide bounds

### Browser Issues
- Clear localStorage if saved layouts cause problems
- Hard refresh (Ctrl+F5) after updates
- Ensure JavaScript is enabled

## Contributing

This project combines visual design tools with AI content generation for automated presentation creation. Feel free to fork and modify for your specific needs.

## License

Personal use project. Modify as needed for your workflows.

## Related Technologies

- **python-pptx**: PowerPoint file generation
- **OpenAI API**: AI content generation
- **Flask**: Web framework
- **HTML5 Canvas**: Visual layout editor