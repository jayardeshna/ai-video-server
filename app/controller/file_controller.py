import os
import re

import pythoncom
from pptx import Presentation
import openai
import spacy
from flask import Blueprint, request, jsonify, current_app
from pptx.util import Inches
from pydub import AudioSegment
from gtts import gTTS
from win32com import client
from moviepy.editor import ImageClip, AudioFileClip, VideoFileClip, concatenate_videoclips
upload_bp = Blueprint('upload', __name__)
nlp = spacy.load('en_core_web_sm')
openai.api_key = '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'


@upload_bp.route('/api/v1/upload', methods=['POST'])
def upload_ppt():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    if file and file.filename.endswith('.pptx'):
        filepath = os.path.join(current_app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filepath)
        slides_data = extract_text_from_ppt(filepath)
        target_language = request.args.get('language', 'Gujarati')
        translated_slides_data = []
        for i, slide in enumerate(slides_data):
            translated_texts = [translate_text(text, target_language) for text in slide['texts']]
            combined_text = " ".join(translated_texts)
            translated_slides_data.append({
                'slide_number': slide['slide_number'],
                'texts': combined_text,
            })
        return jsonify({
            'message': 'File uploaded and text extracted successfully',
            'filename': file.filename,
            'slides': translated_slides_data
        }), 200
    return jsonify({'error': 'Invalid file type'}), 400

@upload_bp.route('/api/v1/generate-video', methods=['POST'])
def generate_video():
    try:
        data = request.get_json()
        slides = data.get('slides', [])
        ppt_path = os.path.join(current_app.config['UPLOAD_FOLDER'], "Automation of Powerpoint to Video Conversion.pptx")
        image_folder = os.path.join(current_app.config['UPLOAD_FOLDER'], 'slides_images')
        slide_images = save_presentation_as_images(ppt_path)
        video_files = []
        for i, slide in enumerate(slides):

            slide_audio_files = []
            audio_files = []

            # Generate one audio file for the combined text
            audio_file = f"slide_{i + 1}.mp3"
            audio_path = os.path.join(current_app.config['UPLOAD_FOLDER'], audio_file)

            text_to_speech(slide.get('texts'), lang='gu', output_file=audio_path)
            slide_audio_files.append(audio_file)
            audio_files.append(audio_file)

            # Create video for each slide
            video_file = f"slide_{i + 1}.mp4"
            video_path = os.path.join(current_app.config['UPLOAD_FOLDER'], video_file)
            create_slide_video(slide_images[i], os.path.join(current_app.config['UPLOAD_FOLDER'], slide_audio_files[0]),
                               video_path)
            video_files.append(video_path)

        print(video_files, "video files")
        # Merge all slide videos into a single video
        clips = [VideoFileClip(video) for video in video_files]
        print(clips, "clips")
        final_clip = concatenate_videoclips(clips)
        final_video_path = os.path.join(current_app.config['UPLOAD_FOLDER'], "final_presentation.mp4")
        final_clip.write_videofile(final_video_path, codec="libx264", audio_codec="aac")

        return jsonify({'message': 'Presentation updated successfully', 'final_video': final_video_path}), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 400


def create_slide_video(image_path, audio_path, output_path):
    image_clip = ImageClip(image_path).set_duration(AudioFileClip(audio_path).duration)
    audio_clip = AudioFileClip(audio_path)
    video = image_clip.set_audio(audio_clip)
    video.write_videofile(output_path, fps=24)

def save_presentation_as_images(ppt_path):
    # Initialize COM library
    pythoncom.CoInitialize()
    pptx_path = os.path.abspath('uploads/Automation of Powerpoint to Video Conversion.pptx')
    output_dir = os.path.abspath('files/output')
    slide_images = []
    os.makedirs(output_dir, exist_ok=True)
    powerpoint = client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = 1
    try:
        presentation = powerpoint.Presentations.Open(pptx_path)
    except Exception as e:
        raise RuntimeError(f"Failed to open PowerPoint presentation: {e}")
    print("2")
    for i, slide in enumerate(presentation.Slides):
        image_path = os.path.join(output_dir, f"slide_{i + 1}.jpg")
        slide_images.append(image_path)
        slide.Export(image_path, "JPG", 1280, 720)
    presentation.Close()
    powerpoint.Quit()
    print(f'All slides have been saved as images in {output_dir}')
    return slide_images


def clean_text(text):
    tags_to_exclude = re.compile(r'/p|/b|/i|/u')  # Add more tags if needed

    text = tags_to_exclude.sub('', text)

    text = re.sub(r'\t+', ' ', text)
    text = re.sub(r'\n+', ' ', text).strip()
    text = re.sub(r'\s+', ' ', text)

    doc = nlp(text)
    cleaned_text = ' '.join(token.text for token in doc)

    return cleaned_text

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
                cleaned_text = clean_text(shape.text)
                slide_content['texts'].append(cleaned_text)

        slides_data.append(slide_content)

    return slides_data


def add_audio_to_slide(ppt_path, slide_number, audio_path):
    # Load the presentation
    prs = Presentation(ppt_path)

    # Ensure the slide number is within the valid range
    if slide_number > len(prs.slides) or slide_number < 1:
        raise ValueError(f"Slide number {slide_number} is out of range. The presentation has {len(prs.slides)} slides.")

    # Select the slide (Note: slide_number - 1 because list index starts from 0)
    slide = prs.slides[slide_number - 1]

    # Check if audio exists and is a valid format
    if not os.path.exists(audio_path) or not audio_path.endswith(('.mp3', '.wav')):
        raise ValueError("Invalid audio file path or format. Only .mp3 or .wav are supported.")

    # Load the audio file using pydub
    audio = AudioSegment.from_file(audio_path)

    # Convert the audio to .wav if it's not already (PowerPoint supports .wav)
    if audio_path.endswith('.mp3'):
        audio_path = audio_path.replace('.mp3', '.wav')
        audio.export(audio_path, format="wav")

    # Add audio to the slide
    slide.shapes.add_movie(audio_path, Inches(0), Inches(0), Inches(1), Inches(1), mime_type='audio/wav')

    # Save the presentation with the new audio
    new_ppt_path = ppt_path.replace('.pptx', '_with_audio.pptx')
    prs.save(new_ppt_path)

    return new_ppt_path


def text_to_speech(text, lang, output_file):
    tts = gTTS(text=text, lang=lang, slow=False)
    tts.save(output_file)
    return output_file