# PowerPoint Layout Designer with AI Content Generation

A comprehensive web-based tool that combines visual PowerPoint slide layout design with AI-powered content generation and image creation. Design custom layouts with templates, generate presentation content with OpenAI, create AI-generated images, and produce professional PowerPoint files with precise formatting and actual images.

## 🚀 Features

### Visual Layout Designer
- **Drag & Drop Interface**: Move elements around the slide canvas with visual feedback
- **Template System**: Choose from Standard Content or Content with Image templates
- **Real-time Preview**: See exactly how your slide will look before generating
- **Element Types**: Title boxes, text boxes, and image placeholders
- **Professional Formatting**: Font selection, sizes, list types, and precise positioning

### AI-Powered Content Generation
- **Topic-Based Planning**: Enter any topic to generate comprehensive presentation outlines
- **Interactive Content Editor**: Modify, add, delete, and reorder slides in the Content Planner
- **Smart Content Creation**: AI generates detailed bullet-point content for each slide
- **Structured Output**: Clean, professional content ready for presentations

### AI Image Generation
- **Automatic Image Creation**: AI generates relevant images for each slide using OpenAI's gpt-image-1 model
- **Smart Prompts**: Automatically creates appropriate image prompts based on slide titles and content
- **Custom Prompts**: Option to provide custom image prompts for specific slides
- **Bulk Generation**: Generate images for multiple slides simultaneously
- **High Quality**: 1024x1024 PNG images optimized for presentations

### Professional PowerPoint Output
- **Actual Images**: Generated AI images are embedded directly into PowerPoint files
- **Proper Bullet Points**: Creates real bullet formatting in PowerPoint (not just text)
- **16:9 Aspect Ratio**: Professional widescreen format (13.333" × 7.5")
- **Font Preservation**: Maintains exact font sizes and formatting
- **Layout Precision**: Your custom layouts are applied consistently across all slides

## 📋 Prerequisites

- Python 3.x
- OpenAI API key with access to gpt-image-1 model
- Modern web browser
- Internet connection for API calls

## 🛠️ Installation

1. **Clone or download the project**
   ```bash
   cd /path/to/your/projects
   git clone <repository-url>
   cd powerpoint-layout-designer
   ```

2. **Install Python dependencies**
   ```bash
   pip install flask openai python-pptx --break-system-packages
   ```

3. **Set up OpenAI API key**
   ```bash
   export OPENAI_API_KEY="your-api-key-here"
   ```

4. **Run the application**
   ```bash
   python3 app.py
   ```

5. **Open your browser**
   ```
   http://localhost:5000
   ```

## 🎯 Complete Workflow

### 1. Design Your Layout
1. **Select Template**: Choose "Standard Content" or "Content with Image" template
2. **Customize Layout**: Use the visual editor to position elements
3. **Configure Formatting**: Set fonts, sizes, and list types
4. **Preview Design**: Toggle between Title and Content slide views

### 2. Plan Your Content
1. **Enter Topic**: Type your presentation topic in the sidebar
2. **Generate Draft**: Click "Generate Draft" for AI-created slide outline
3. **Open Content Planner**: Review and edit the generated structure
4. **Refine Structure**: Add, remove, or reorder slides as needed

### 3. Generate Content & Images
1. **Create Content**: Click "Generate Content" for detailed slide text
2. **Generate Images**: Click "Generate Images" to create AI images for all slides
3. **Review Results**: Check generated content and images
4. **Make Edits**: Modify any content or regenerate specific images

### 4. Create Final Presentation
1. **Download PowerPoint**: Click "Create PowerPoint" for the final PPTX file
2. **Complete Package**: Get professional slides with custom layouts, AI content, and actual images
3. **Ready to Present**: Open in PowerPoint for final touches or immediate use

## 🖼️ Image Generation Features

### Automatic Image Creation
- **Smart Prompts**: AI analyzes slide titles and content to create relevant image prompts
- **Educational Focus**: Images designed specifically for presentation contexts
- **Consistent Quality**: 1024x1024 resolution for crisp display

### Bulk Generation
- **Multiple Slides**: Generate images for all slides at once using concurrent processing
- **Progress Tracking**: Real-time feedback on generation progress
- **Error Handling**: Graceful fallbacks for failed generations

### Custom Control
- **Custom Prompts**: Provide your own image prompts for specific slides
- **Regeneration**: Re-generate individual images with different prompts
- **Preview Integration**: Images appear immediately in the web interface

## 📐 Templates

### Standard Content Template
- Full-width title at top
- Single-column content area
- Ideal for text-heavy presentations
- Clean, professional layout

### Content with Image Template
- Full-width title across top
- Two-column layout below
- Text content on left with bullet points
- Image area on right side
- Perfect for visual presentations

## 🔧 Technical Details

### AI Integration
- **Content Generation**: OpenAI GPT-3.5-turbo for structured text content
- **Image Generation**: OpenAI gpt-image-1 for high-quality presentation images
- **Smart Prompting**: Optimized prompts for educational and business contexts

### Image Processing
- **File Management**: Automatic saving and organization of generated images
- **Format Handling**: PNG format with base64 decoding
- **PowerPoint Integration**: Direct embedding of images into PPTX files

### Layout System
- **Coordinate Mapping**: Precise conversion from web canvas to PowerPoint inches
- **Template Engine**: Flexible system supporting multiple layout types
- **Element Positioning**: Exact positioning with drag-and-drop interface

## 📁 Project Structure

```
powerpoint-layout-designer/
├── app.py                          # Flask application with AI integration
├── templates/
│   └── index.html                  # Web interface with image generation
├── static/
│   └── generated_images/           # AI-generated images storage
├── gitignore/                      # Large files (excluded from git)
└── README.md                       # This documentation
```

## 🔌 API Endpoints

### Core Features
- `GET /` - Main application interface
- `POST /generate_draft` - Create slide outline from topic
- `POST /generate_content` - Generate detailed slide content
- `POST /create_presentation` - Build final PPTX with images

### Image Generation
- `POST /generate_image` - Generate single image for a slide
- `POST /generate_images_bulk` - Generate images for multiple slides
- `POST /generate_image_prompt` - Create optimized image prompt
- `GET /static/generated_images/<filename>` - Serve generated images

### Development Tools
- `POST /generate_code` - Export python-pptx code
- `POST /download_code` - Download generated code

## 🎨 Image Features in Detail

### Generated Image Integration
- **Seamless Embedding**: AI-generated images automatically appear in downloaded PowerPoint files
- **Smart Placement**: Images positioned precisely according to your layout design
- **Fallback System**: Graceful handling when images can't be generated or loaded

### Image Management
- **Automatic Storage**: Generated images saved in `static/generated_images/`
- **Unique Naming**: UUID-based filenames prevent conflicts
- **Web Serving**: Flask serves images for preview and PowerPoint embedding

## 🐛 Troubleshooting

### Image Generation Issues
**Images not generating:**
- Verify OpenAI API key has access to gpt-image-1 model
- Check internet connection for API calls
- Review browser console for error messages

**Images not appearing in PowerPoint:**
- This issue has been fixed - images now embed properly
- Check that image files exist in `static/generated_images/`
- Verify PowerPoint download includes actual images, not placeholders

### API and Performance
**Slow image generation:**
- Image generation uses concurrent processing for multiple slides
- Individual images may take 10-15 seconds each
- Large presentations may take several minutes for full image generation

**OpenAI API errors:**
- Ensure API key is valid and has sufficient credits
- Check API rate limits for high-volume usage
- Verify model access permissions

## 📝 Example Use Cases

### Business Presentations
1. **Topic**: "Q4 Sales Strategy"
2. **AI generates**: Executive summary, market analysis, strategy slides, implementation timeline
3. **Images**: Charts, business graphics, strategy diagrams
4. **Result**: Professional presentation ready for stakeholder meetings

### Educational Content
1. **Topic**: "Introduction to Photosynthesis"
2. **AI generates**: Process overview, detailed steps, environmental impact, applications
3. **Images**: Plant diagrams, process illustrations, scientific visuals
4. **Result**: Engaging educational slides with relevant imagery

### Project Planning
1. **Topic**: "Website Redesign Project"
2. **AI generates**: Project scope, timeline, resource requirements, deliverables
3. **Images**: Wireframes, design mockups, process flows
4. **Result**: Complete project presentation with visual aids

## 🔮 Recent Updates

### Version 2.0 - AI Image Integration
- **NEW**: Full AI image generation using OpenAI gpt-image-1
- **NEW**: Bulk image generation for multiple slides
- **NEW**: Custom prompt support for specific image requirements
- **FIXED**: Images now properly embed in downloaded PowerPoint files
- **IMPROVED**: Enhanced error handling and fallback systems

## 🤝 Contributing

This project represents a complete presentation creation pipeline combining:
- Visual design tools for layout customization
- AI content generation for structured text
- AI image generation for visual enhancement
- Professional PowerPoint output with embedded media

Feel free to extend functionality, add new templates, or integrate additional AI models.

## 📄 License

Personal use project. Modify and extend as needed for your presentation workflows.

## 🔗 Related Technologies

- **Flask**: Web framework for the visual editor
- **OpenAI API**: GPT-3.5-turbo for content, gpt-image-1 for images
- **python-pptx**: PowerPoint file generation with precise formatting
- **HTML5 Canvas**: Interactive layout designer
- **Concurrent Processing**: Multi-threaded image generation for performance