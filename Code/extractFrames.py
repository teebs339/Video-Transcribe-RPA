import cv2
import os
import pandas as pd
import sys

def parse_srt_time(srt_time):
    parts = srt_time.split(':')
    hours, minutes = map(float, parts[:2])  # First two are hours and minutes
    seconds_parts = parts[2].split(',')  # Split the seconds and milliseconds
    
    seconds = float(seconds_parts[0])  # Get the seconds
    milliseconds = float(seconds_parts[1]) / 1000 if len(seconds_parts) > 1 else 0  # Convert milliseconds to seconds if present
    
    return hours * 3600 + minutes * 60 + seconds + milliseconds

def get_frame_interval(time_in_video):
    if time_in_video <= 3:
        return 0.25  # First 3 seconds
    elif time_in_video <= 15:
        return 0.5  # 3 to 15 seconds
    else:
        return 0.75  # After 15 seconds

def extract_frames(video_path, output_folder, interval, timestamps):
    vidcap = cv2.VideoCapture(video_path)
    fps = vidcap.get(cv2.CAP_PROP_FPS)
    frame_count = 0
    saved_count = 0

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    success, image = vidcap.read()
    while success:
        time_in_video = frame_count / fps
        transcript = "frame"
        
        for index, row in timestamps.iterrows():
            start_time = parse_srt_time(row['Start time'])
            end_time = parse_srt_time(row['End time'])
            
            if start_time <= time_in_video <= end_time:
                transcript = row['Transcript']  # Use the transcript from timestamps
                break
        
        frame_interval = get_frame_interval(time_in_video)  # Determine the interval based on time
        
        # Save a frame according to the calculated interval
        if time_in_video >= frame_interval * saved_count:
            frame_filename = f"{saved_count + 1} - {transcript}_{saved_count:05d}.jpg"
            output_path = os.path.join(output_folder, frame_filename)
            cv2.imwrite(output_path, image)
            saved_count += 1
            print(f"Saved {output_path}")
        
        success, image = vidcap.read()
        frame_count += 1

    vidcap.release()
    print(f"Extracted {saved_count} frames to {output_folder}.")

if __name__ == "__main__":
    user_profile_directory = os.path.expanduser("~")
    excel_file_path = os.path.join(user_profile_directory, "pythonTemp.xlsx")
    timestamps_file_path = os.path.join(user_profile_directory, "Timestamps.xlsx")

    df = pd.read_excel(excel_file_path, sheet_name='Sheet1')
    video_path = df.loc[df['Name'] == 'pathVideo', 'Value'].values[0]
    output_folder = df.loc[df['Name'] == 'pathFolder', 'Value'].values[0]
    interval = float(df.loc[df['Name'] == 'interval', 'Value'].values[0])  # Ignored in logic but passed

    timestamps = pd.read_excel(timestamps_file_path, sheet_name='Sheet1')

    extract_frames(video_path, output_folder, interval, timestamps)
    print("Frames capture - Complete")
