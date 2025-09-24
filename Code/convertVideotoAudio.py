import os
import pandas as pd
from moviepy.editor import VideoFileClip

def get_paths_from_excel(excel_file_path):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(excel_file_path)
    
    # Extract video and output paths based on the "Name" column
    video_path = df.loc[df['Name'] == 'Input path', 'Value'].values[0]
    output_folder = df.loc[df['Name'] == 'Output path', 'Value'].values[0]

    return video_path, output_folder

def convert_video_to_audio(video_path, audio_path):
    try:
        # Load the video file
        video_clip = VideoFileClip(video_path)
        
        # Extract the audio and write it to a file
        audio_clip = video_clip.audio
        audio_clip.write_audiofile(audio_path)

        # Close the clips
        audio_clip.close()
        video_clip.close()

        print(f"Audio file saved successfully as: {audio_path}")
    except Exception as e:
        print(f"An error occurred: {e}")

# Example usage
user_profile_directory = os.path.expanduser("~")  # Get the user profile directory
excel_file_path = os.path.join(user_profile_directory, "pythonTemp.xlsx")  # Path to the Excel file in user profile

# Get the video path and output folder from the Excel file
video_file, output_folder = get_paths_from_excel(excel_file_path)

# Set the audio file path
audio_file = os.path.join(output_folder, "ConvertedAudio.mp3")

# Convert the video to audio
convert_video_to_audio(video_file, audio_file)
