import os
import re
import shutil
import pythoncom
from win32com.client import Dispatch
import comtypes
from pptx import Presentation
import openai
from flask import Blueprint, request, jsonify, current_app
from pptx.util import Inches
from pydub import AudioSegment
from gtts import gTTS
from comtypes.client import CreateObject
from moviepy.editor import ImageClip, AudioFileClip, VideoFileClip, concatenate_videoclips
upload_bp = Blueprint('upload', __name__)
# nlp = spacy.load('en_core_web_sm')
openai.api_key = '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'


@upload_bp.route('/api/v1/upload', methods=['POST'])
def upload_ppt():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    upload_folder = current_app.config['PPT_FOLDER']

    # Check if the directory exists, and create it if it doesn't
    if not os.path.exists(upload_folder):
        os.makedirs(upload_folder)

    if file and file.filename.endswith('.pptx'):
        filepath = os.path.join(current_app.config['PPT_FOLDER'], file.filename)
        file.save(filepath)
        translated_filepath = os.path.join(current_app.config['PPT_FOLDER'], 'Translated_' + file.filename)
        target_language = request.args.get('language', 'Gujarati')
        translate_ppt(filepath, translated_filepath, target_language)
        slides_data = extract_text_from_ppt(translated_filepath)

        translated_slides_data = []
        for slide in slides_data:
            combined_text = " ".join(slide['texts'])
            translated_slides_data.append({
                'slide_number': slide['slide_number'],
                'texts': combined_text,
            })
        return jsonify({
            'message': 'File uploaded and text extracted successfully',
            'filename': 'Translated_' + file.filename,
            'slides': translated_slides_data
        }), 200
    return jsonify({'error': 'Invalid file type'}), 400

@upload_bp.route('/api/v1/generate-video', methods=['POST'])
def generate_video():
    try:
        data = request.get_json()
        slides = data.get('slides', [])
        file_name = data.get('filename')
        ppt_path = os.path.join(current_app.config['PPT_FOLDER'], file_name)
        image_folder = os.path.join(current_app.config['IMAGE_FOLDER'])
        slide_images = save_presentation_as_images(ppt_path, image_folder)
        video_files = []
        for i, slide in enumerate(slides):
            slide_audio_files = []
            audio_files = []
            # Generate one audio file for the combined text
            audio_file = f"slide_{i + 1}.mp3"
            audio_folder = current_app.config['AUDIO_FOLDER']
            if not os.path.exists(audio_folder):
                os.makedirs(audio_folder)
            audio_path = os.path.join(current_app.config['AUDIO_FOLDER'], audio_file)
            text_to_speech(slide.get('texts'), lang='gu', output_file=audio_path)
            slide_audio_files.append(audio_file)
            audio_files.append(audio_file)
            video_file = f"slide_{i + 1}.mp4"
            video_folder = current_app.config['VIDEO_PATH']
            if not os.path.exists(video_folder):
                os.makedirs(video_folder)
            video_path = os.path.join(current_app.config['VIDEO_PATH'], video_file)
            create_slide_video(slide_images[i], os.path.join(current_app.config['AUDIO_FOLDER'], slide_audio_files[0]),
                               video_path)
            video_files.append(video_path)

        # Merge all slide videos into a single video
        clips = [VideoFileClip(video) for video in video_files]
        final_clip = concatenate_videoclips(clips)
        final_video_path = os.path.join(current_app.config['UPLOAD_FOLDER'], "final_presentation.mp4")
        final_clip.write_videofile(final_video_path, codec="libx264", audio_codec="aac")
        shutil.rmtree(current_app.config['AUDIO_FOLDER'], ignore_errors=True)
        shutil.rmtree(current_app.config['IMAGE_FOLDER'], ignore_errors=True)
        shutil.rmtree(current_app.config['PPT_FOLDER'], ignore_errors=True)
        shutil.rmtree(current_app.config['VIDEO_PATH'], ignore_errors=True)
        return jsonify({'message': 'Presentation updated successfully', 'final_video': os.path.abspath(final_video_path)}), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 400

def translate_ppt(input_ppt_path, output_ppt_path, target_language="Hindi"):
    # Load the original presentation
    prs = Presentation(input_ppt_path)

    # Iterate through slides and shapes to translate the text while maintaining the format
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    original_text = paragraph.text
                    if original_text.strip():
                        translated_text = translate_text(original_text, target_language)

                        # Replace the text with the translation
                        original_font = paragraph.font
                        paragraph.text = translated_text

                        # Restore the original font settings
                        paragraph.font.name = original_font.name
                        paragraph.font.size = original_font.size
                        paragraph.font.bold = original_font.bold
                        paragraph.font.italic = original_font.italic

    # Save the translated presentation
    prs.save(output_ppt_path)
    print(f"Translated PowerPoint saved as {output_ppt_path}")


def create_slide_video(image_path, audio_path, output_path):
    image_clip = ImageClip(image_path).set_duration(AudioFileClip(audio_path).duration)
    audio_clip = AudioFileClip(audio_path)

    video = image_clip.set_audio(audio_clip)

    video.write_videofile(output_path, fps=24)

def save_presentation_as_images(ppt_path, output_folder):
    pythoncom.CoInitialize()
    powerpoint = CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1
    slide_images = []
    ppt_abs_path = os.path.abspath(ppt_path)
    output_folder_abs = os.path.abspath(output_folder)
    # Open the presentation
    presentation = powerpoint.Presentations.Open(ppt_abs_path)

    if not os.path.exists(output_folder_abs):
        os.makedirs(output_folder_abs)

    for i, slide in enumerate(presentation.Slides):
        slide_image_path = os.path.join(output_folder_abs, f"slide_{i + 1}.jpg")
        slide_images.append(slide_image_path)
        slide.Export(slide_image_path, "JPG", 1920, 1080)  # You can adjust resolution here
        print(f"Saved {slide_image_path}")

    # Close the presentation and quit PowerPoint
    presentation.Close()
    powerpoint.Quit()
    pythoncom.CoUninitialize()
    return slide_images


def translate_text(text, target_language):
    prompt = f"Translate the following text into {target_language}:\n\n{text}"
    response = openai.ChatCompletion.create(
        model="gpt-4",
        messages=[{"role": "system", "content": "You are a helpful assistant."},
                  {"role": "user", "content": prompt}]
    )
    translated_text = response.choices[0].message['content']
    return translated_text


def extract_text_from_ppt(filepath):
    prs = Presentation(filepath)
    slides_data = []

    for slide_number, slide in enumerate(prs.slides, start=1):
        slide_content = {
            'slide_number': slide_number,
            'texts': []
        }

        for shape in slide.shapes:
            if hasattr(shape, "text"):
                # cleaned_text = clean_text(shape.text)
                slide_content['texts'].append(shape.text)

        slides_data.append(slide_content)

    return slides_data


def text_to_speech(text, lang, output_file):
    tts = gTTS(text=text, lang=lang, slow=False)
    tts.save(output_file)
    return output_file