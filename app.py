from flask import Flask, render_template, request, jsonify, send_file
import json
import io
import os
from datetime import datetime
import openai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.enum.text import MSO_ANCHOR
from pptx.dml.color import RGBColor
import re
import math

app = Flask(__name__)

# Initialize OpenAI client (users should set OPENAI_API_KEY environment variable)
openai_client = openai.OpenAI(api_key=os.getenv('OPENAI_API_KEY'))


def parse_bullet_points(content):
    """Parse content and extract bullet points - each line becomes a bullet point"""
    if not content:
        return []
    
    lines = content.strip().split('\n')
    bullet_points = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        # Remove any bullet symbols that might be present (including hyphens)
        if line.startswith(('- ', '• ', '* ', '· ', '○ ', '▪ ', '‣ ')):
            bullet_text = line[2:].strip()
            if bullet_text:
                bullet_points.append(bullet_text)
        elif line.startswith('-'):
            # Handle cases where there's just a dash without space
            bullet_text = line[1:].strip()
            if bullet_text:
                bullet_points.append(bullet_text)
        elif line and not any(line.lower().startswith(word) for word in ['slide', 'title:', 'content:']):
            # Include non-empty lines that aren't headers - each line becomes a bullet
            bullet_points.append(line)
    
    return bullet_points


@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate_code', methods=['POST'])
def generate_code():
    layout_data = request.json
    
    # Generate python-pptx code
    code = generate_pptx_code(layout_data)
    
    # Create downloadable file
    file_content = io.StringIO(code)
    file_content.seek(0)
    
    return jsonify({
        'code': code,
        'filename': f'powerpoint_layout_{datetime.now().strftime("%Y%m%d_%H%M%S")}.py'
    })

@app.route('/download_code', methods=['POST'])
def download_code():
    code = request.json.get('code')
    filename = request.json.get('filename', 'powerpoint_layout.py')
    
    # Create file in memory
    file_buffer = io.BytesIO()
    file_buffer.write(code.encode('utf-8'))
    file_buffer.seek(0)
    
    return send_file(
        file_buffer,
        as_attachment=True,
        download_name=filename,
        mimetype='text/plain'
    )

def generate_pptx_code(layout_data):
    """Generate python-pptx code from layout data"""
    
    code_lines = [
        "from pptx import Presentation",
        "from pptx.util import Inches, Pt",
        "from pptx.enum.text import PP_ALIGN",
        "from pptx.enum.text import MSO_ANCHOR",
        "from pptx.dml.color import RGBColor",
        "",
        "# Create presentation",
        "pres = Presentation()",
        ""
    ]
    
    slide_type = layout_data.get('slide_type', 'content')
    elements = layout_data.get('elements', [])
    
    if slide_type == 'title':
        code_lines.append("# Add title slide")
        code_lines.append("slide_layout = pres.slide_layouts[6]  # Blank layout")
    else:
        code_lines.append("# Add content slide") 
        code_lines.append("slide_layout = pres.slide_layouts[6]  # Blank layout")
    
    code_lines.append("slide = pres.slides.add_slide(slide_layout)")
    code_lines.append("")
    
    # Add elements
    for i, element in enumerate(elements):
        if element['type'] in ['textbox', 'title']:
            code_lines.extend(generate_textbox_code(element, i))
        elif element['type'] == 'image':
            code_lines.extend(generate_image_code(element, i))
    
    code_lines.extend([
        "",
        "# Save presentation",
        "pres.save('generated_presentation.pptx')",
        "print('Presentation saved as generated_presentation.pptx')"
    ])
    
    return "\n".join(code_lines)

def generate_textbox_code(element, index):
    """Generate code for text box element"""
    
    left = element['left']
    top = element['top'] 
    width = element['width']
    height = element['height']
    font_size = element.get('font_size', 18)
    font_name = element.get('font_name', 'Arial')
    list_type = element.get('list_type', 'none')
    content = element.get('content', 'Sample text')
    
    lines = [
        f"# Add {'title' if element['type'] == 'title' else 'text'} box {index + 1}",
        f"textbox_{index} = slide.shapes.add_textbox(",
        f"    Inches({left}), Inches({top}), Inches({width}), Inches({height})",
        f")",
        f"text_frame_{index} = textbox_{index}.text_frame",
        f"text_frame_{index}.clear()  # Clear default paragraph",
        f"text_frame_{index}.word_wrap = True",
        ""
    ]
    
    # Check if content has bullet points
    if list_type == 'bullet' and ('\\n' in content or '\n' in content):
        bullet_points = parse_bullet_points(content)
        lines.extend([
            f"# Add bullet points",
            f"bullet_items = {repr(bullet_points)}",
            f"for i, item in enumerate(bullet_items):",
            f"    if item.strip():",
            f"        if i == 0:",
            f"            p = text_frame_{index}.paragraphs[0]",
            f"        else:",
            f"            p = text_frame_{index}.add_paragraph()",
            f"        p.text = item.strip()",
            f"        p.level = 0  # First level bullet",
            f"        # Format the paragraph",
            f"        for run in p.runs:",
            f"            font = run.font",
            f"            font.name = '{font_name}'",
            f"            font.size = Pt({font_size})",
        ])
    else:
        # Simple text without bullets
        lines.extend([
            f"p = text_frame_{index}.paragraphs[0]",
            f"p.text = '{content}'",
            f"",
            f"# Format text",
            f"for run in p.runs:",
            f"    font = run.font",
            f"    font.name = '{font_name}'",
            f"    font.size = Pt({font_size})",
        ])
        
        # Add alignment for title
        if element['type'] == 'title':
            lines.insert(-5, "p.alignment = PP_ALIGN.CENTER")
    
    lines.append("")
    
    return lines

def generate_image_code(element, index):
    """Generate code for image element"""
    
    left = element['left']
    top = element['top']
    width = element['width'] 
    height = element['height']
    
    lines = [
        f"# Add image {index + 1}",
        f"# Note: Replace 'path_to_image.jpg' with your actual image path",
        f"image_{index} = slide.shapes.add_picture(",
        f"    'path_to_image.jpg',",
        f"    Inches({left}), Inches({top}), Inches({width}), Inches({height})",
        f")",
        ""
    ]
    
    return lines

@app.route('/generate_draft', methods=['POST'])
def generate_draft():
    """Generate a draft outline of slide titles based on user topic"""
    data = request.json
    topic = data.get('topic', '')
    
    if not topic:
        return jsonify({'error': 'Topic is required'}), 400
    
    try:
        # Generate slide outline using OpenAI
        response = openai_client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are a presentation expert. Create a clear, logical outline for a PowerPoint presentation. The first line should be the main presentation title (what goes on the title slide). The following lines should be content slide titles. Return only slide titles, one per line, without numbering, bullet points, or prefixes like 'Title:' or 'Slide 1:'. Provide 5-9 total lines."},
                {"role": "user", "content": f"Create a presentation outline about: {topic}\n\nExample format:\nThe Complete Guide to Solar Energy\nIntroduction to Solar Technology\nTypes of Solar Panels\nInstallation Process\nCost and Benefits\nMaintenance and Care"}
            ],
            max_tokens=300,
            temperature=0.7
        )
        
        slides_text = response.choices[0].message.content.strip()
        slides = [line.strip() for line in slides_text.split('\n') if line.strip()]
        
        # Format slides with types and clean up titles
        formatted_slides = []
        
        # First slide is always the title slide using the first line from OpenAI
        if slides:
            title_slide_title = slides[0]
            # Remove common prefixes
            title_slide_title = re.sub(r'^Slide\s*\d+\s*:\s*', '', title_slide_title, flags=re.IGNORECASE)
            title_slide_title = re.sub(r'^\d+\.\s*', '', title_slide_title)
            
            formatted_slides.append({
                'id': 0,
                'title': title_slide_title,
                'type': 'title',
                'content': ''  # Title slides typically have no content or just subtitle
            })
            
            # Remaining slides are content slides
            for i, title in enumerate(slides[1:], 1):
                # Remove common prefixes
                cleaned_title = re.sub(r'^Slide\s*\d+\s*:\s*', '', title, flags=re.IGNORECASE)
                cleaned_title = re.sub(r'^\d+\.\s*', '', cleaned_title)
                
                formatted_slides.append({
                    'id': i,
                    'title': cleaned_title,
                    'type': 'content',
                    'content': 'Content will be generated after approval'
                })
        
        return jsonify({
            'slides': formatted_slides,
            'topic': topic
        })
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/generate_content', methods=['POST'])
def generate_content():
    """Generate full content for approved slide outline"""
    data = request.json
    slides = data.get('slides', [])
    topic = data.get('topic', '')
    content_layout = data.get('content_layout', {})
    
    if not slides:
        return jsonify({'error': 'Slides are required'}), 400
    
    
    try:
        # Generate content for each slide
        for slide in slides:
            if slide['type'] == 'content' and not slide.get('content_generated', False):
                response = openai_client.chat.completions.create(
                    model="gpt-3.5-turbo",
                    messages=[
                        {"role": "system", "content": "You are a presentation content writer. Create bullet points for PowerPoint slides. Provide each point as a separate line with NO bullet symbols, NO dashes, NO prefixes - just plain text. Each line will automatically become a bullet point in the presentation. Be concise and impactful."},
                        {"role": "user", "content": f"Create content for this slide about '{topic}':\nSlide title: {slide['title']}\n\nProvide 3-5 bullet points. Each point should be on its own line with no bullet symbols. Example format:\nPoint one text here\nPoint two text here\nPoint three text here"}
                    ],
                    max_tokens=200,
                    temperature=0.7
                )
                
                slide['content'] = response.choices[0].message.content.strip()
                slide['content_generated'] = True
        
        return jsonify({'slides': slides})
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/create_presentation', methods=['POST'])
def create_presentation():
    """Create final PPTX file with generated content and custom layouts"""
    data = request.json
    slides_data = data.get('slides', [])
    title_layout = data.get('title_layout', {})
    content_layout = data.get('content_layout', {})
    
    if not slides_data:
        return jsonify({'error': 'Slides data is required'}), 400
    
    # Debug: print the received layouts
    print("Title layout received:", title_layout)
    print("Content layout received:", content_layout)
    
    try:
        # Create presentation with 16:9 aspect ratio
        pres = Presentation()
        
        # Set slide size to 16:9 widescreen format
        pres.slide_width = Inches(13.333)  # 16:9 ratio width
        pres.slide_height = Inches(7.5)    # 16:9 ratio height
        
        for slide_data in slides_data:
            slide_type = slide_data['type']
            
            # Use the appropriate layout based on slide type
            if slide_type == 'title':
                layout_config = title_layout
            else:
                layout_config = content_layout
            
            print(f"Using layout for {slide_type} slide:", layout_config)
            
            # Add blank slide
            slide_layout = pres.slide_layouts[6]  # Blank layout
            slide = pres.slides.add_slide(slide_layout)
            
            # Add elements based on layout configuration
            elements = layout_config.get('elements', [])
            print(f"Processing {len(elements)} elements for slide: {slide_data['title']}")
            
            for element in elements:
                print(f"Processing element: {element}")
                
                if element['type'] in ['textbox', 'title']:
                    # Add text box with EXACT positioning from the designer
                    textbox = slide.shapes.add_textbox(
                        Inches(element['left']), 
                        Inches(element['top']), 
                        Inches(element['width']), 
                        Inches(element['height'])
                    )
                    text_frame = textbox.text_frame
                    text_frame.clear()  # Clear default content
                    
                    # Configure text frame properties
                    text_frame.word_wrap = True
                    text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
                    
                    # Determine what content to use
                    if element['type'] == 'title':
                        # For title elements, always use the slide title
                        content_text = slide_data['title']
                        p = text_frame.paragraphs[0]
                        p.text = content_text
                        p.alignment = PP_ALIGN.CENTER
                        
                        # Apply title formatting from the designer
                        for run in p.runs:
                            font = run.font
                            font.name = element.get('font_name', 'Calibri')
                            font.size = Pt(element.get('font_size', 28))
                            font.bold = True
                    else:
                        # For textbox elements, use slide content
                        content_text = slide_data.get('content', '')
                        
                        if slide_type == 'title':
                            # For title slides, textbox shows subtitle/description
                            p = text_frame.paragraphs[0]
                            p.text = content_text if content_text else ""
                            p.alignment = PP_ALIGN.CENTER
                            
                            if p.runs:
                                for run in p.runs:
                                    font = run.font
                                    font.name = element.get('font_name', 'Calibri')
                                    font.size = Pt(element.get('font_size', 18))
                        else:
                            # For content slides, parse and add bullet points if list_type is bullet
                            if element.get('list_type', 'none') == 'bullet':
                                bullet_points = parse_bullet_points(content_text)
                                print(f"Parsed bullet points: {bullet_points}")
                                
                                if bullet_points:
                                    # Add bullet points - each line becomes a bullet
                                    for i, bullet_text in enumerate(bullet_points):
                                        if i == 0:
                                            p = text_frame.paragraphs[0]
                                        else:
                                            p = text_frame.add_paragraph()
                                        
                                        p.text = bullet_text.strip()
                                        p.level = 0  # First level bullet
                                        
                                        # Apply formatting from the designer
                                        # Format the paragraph first, then apply to runs
                                        font = p.font
                                        font.name = element.get('font_name', 'Calibri')
                                        font.size = Pt(element.get('font_size', 18))
                                        
                                        # Also apply to any existing runs
                                        for run in p.runs:
                                            run_font = run.font
                                            run_font.name = element.get('font_name', 'Calibri')
                                            run_font.size = Pt(element.get('font_size', 18))
                                else:
                                    # No bullet points, just add the text
                                    p = text_frame.paragraphs[0]
                                    p.text = content_text
                                    
                                    if p.runs:
                                        for run in p.runs:
                                            font = run.font
                                            font.name = element.get('font_name', 'Calibri')
                                            font.size = Pt(element.get('font_size', 18))
                            else:
                                # Plain text without bullets
                                p = text_frame.paragraphs[0]
                                p.text = content_text
                                
                                if p.runs:
                                    for run in p.runs:
                                        font = run.font
                                        font.name = element.get('font_name', 'Calibri')
                                        font.size = Pt(element.get('font_size', 18))
                                        
                elif element['type'] == 'image':
                    # Add image placeholder with exact positioning
                    textbox = slide.shapes.add_textbox(
                        Inches(element['left']), 
                        Inches(element['top']), 
                        Inches(element['width']), 
                        Inches(element['height'])
                    )
                    text_frame = textbox.text_frame
                    text_frame.clear()
                    p = text_frame.paragraphs[0]
                    p.text = "[IMAGE PLACEHOLDER]"
                    p.alignment = PP_ALIGN.CENTER
        
        # Save presentation to memory
        file_buffer = io.BytesIO()
        pres.save(file_buffer)
        file_buffer.seek(0)
        
        filename = f"presentation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
        
        return send_file(
            file_buffer,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
        
    except Exception as e:
        print("Error creating presentation:", str(e))
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)
