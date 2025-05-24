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
    """Parse content and extract bullet points, removing any bullet symbols"""
    lines = content.strip().split('\n')
    bullet_points = []
    
    for line in lines:
        line = line.strip()
        # Remove any bullet symbols that might be present
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
            # Include non-empty lines that aren't headers
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
    if list_type == 'bullet' and '\\n' in content:
        lines.extend([
            f"# Add bullet points",
            f"bullet_items = {repr(content.split('\\n'))}",
            f"for i, item in enumerate(bullet_items):",
            f"    if item.strip():",
            f"        if i == 0:",
            f"            p = text_frame_{index}.paragraphs[0]",
            f"        else:",
            f"            p = text_frame_{index}.add_paragraph()",
            f"        p.text = item.strip().lstrip('- •*·○▪‣')",
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
                {"role": "system", "content": "You are a presentation expert. Create a clear, logical outline for a PowerPoint presentation. Return only slide titles, one per line, without numbering or bullet points. Include a title slide and 4-8 content slides."},
                {"role": "user", "content": f"Create a presentation outline about: {topic}"}
            ],
            max_tokens=300,
            temperature=0.7
        )
        
        slides_text = response.choices[0].message.content.strip()
        slides = [line.strip() for line in slides_text.split('\n') if line.strip()]
        
        # Format slides with types and clean up titles
        formatted_slides = []
        for i, title in enumerate(slides):
            # Remove common prefixes like "Slide 1:", "Slide 2:", etc.
            cleaned_title = re.sub(r'^Slide\s*\d+\s*:\s*', '', title, flags=re.IGNORECASE)
            cleaned_title = re.sub(r'^\d+\.\s*', '', cleaned_title)  # Remove numbering like "1. "
            
            slide_type = 'title' if i == 0 else 'content'
            formatted_slides.append({
                'id': i,
                'title': cleaned_title,
                'type': slide_type,
                'content': '' if slide_type == 'title' else 'Content will be generated after approval'
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
                        {"role": "system", "content": "You are a presentation content writer. Create bullet points for PowerPoint slides. Provide each point as a separate line without any bullet symbols or dashes. Be concise and impactful."},
                        {"role": "user", "content": f"Create content for this slide about '{topic}':\nSlide title: {slide['title']}\n\nProvide 3-5 bullet points. Each bullet should be concise and fit on 1-2 lines."}
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
    
    try:
        # Create presentation
        pres = Presentation()
        
        for slide_data in slides_data:
            slide_type = slide_data['type']
            layout_config = title_layout if slide_type == 'title' else content_layout
            
            # Add blank slide
            slide_layout = pres.slide_layouts[6]  # Blank layout
            slide = pres.slides.add_slide(slide_layout)
            
            # Add elements based on layout configuration
            elements = layout_config.get('elements', [])
            
            for element in elements:
                if element['type'] in ['textbox', 'title']:
                    # Add text box
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
                    
                    # Set content based on element type and slide data
                    if element['type'] == 'title':
                        # For title elements, use the slide title
                        p = text_frame.paragraphs[0]
                        p.text = slide_data['title']
                        p.alignment = PP_ALIGN.CENTER
                        
                        # Apply title formatting
                        for run in p.runs:
                            font = run.font
                            font.name = element.get('font_name', 'Arial')
                            font.size = Pt(element.get('font_size', 28))
                            font.bold = True
                    else:
                        # For content elements
                        content = slide_data.get('content', '')
                        
                        if slide_type == 'title':
                            # For title slides, add subtitle or description
                            p = text_frame.paragraphs[0]
                            p.text = content if content else f"Presentation on {slide_data['title']}"
                            p.alignment = PP_ALIGN.CENTER
                            
                            for run in p.runs:
                                font = run.font
                                font.name = element.get('font_name', 'Arial')
                                font.size = Pt(element.get('font_size', 18))
                        else:
                            # For content slides, parse and add bullet points
                            bullet_points = parse_bullet_points(content)
                            
                            if bullet_points and element.get('list_type', 'bullet') == 'bullet':
                                # Add bullet points
                                for i, bullet_text in enumerate(bullet_points):
                                    if i == 0:
                                        p = text_frame.paragraphs[0]
                                    else:
                                        p = text_frame.add_paragraph()
                                    
                                    p.text = bullet_text
                                    p.level = 0  # First level bullet
                                    
                                    # Apply formatting
                                    for run in p.runs:
                                        font = run.font
                                        font.name = element.get('font_name', 'Arial')
                                        font.size = Pt(element.get('font_size', 18))
                            else:
                                # Plain text without bullets
                                p = text_frame.paragraphs[0]
                                p.text = content
                                
                                for run in p.runs:
                                    font = run.font
                                    font.name = element.get('font_name', 'Arial')
                                    font.size = Pt(element.get('font_size', 18))
        
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
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)
