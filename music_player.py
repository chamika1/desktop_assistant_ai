import pygame
from youtube_dl import YoutubeDL
import urllib.request
import re

def get_youtube_url(song_name):
    # Search for the song on YouTube
    search_query = urllib.parse.quote(song_name)
    html = urllib.request.urlopen(f"https://www.youtube.com/results?search_query={search_query}")
    video_ids = re.findall(r"watch\?v=(\S{11})", html.read().decode())
    if video_ids:
        return f"https://www.youtube.com/watch?v={video_ids[0]}"
    return None

def play_song(song_name):
    try:
        # Get YouTube URL for the song
        url = get_youtube_url(song_name)
        if not url:
            return False, "Could not find the song"

        # Configure youtube-dl
        ydl_opts = {
            'format': 'bestaudio/best',
            'postprocessors': [{
                'key': 'FFmpegExtractAudio',
                'preferredcodec': 'mp3',
                'preferredquality': '192',
            }],
            'outtmpl': 'temp_song.%(ext)s'
        }

        # Download the song
        with YoutubeDL(ydl_opts) as ydl:
            ydl.download([url])

        # Initialize pygame mixer
        pygame.mixer.init()
        pygame.mixer.music.load('temp_song.mp3')
        pygame.mixer.music.play()
        
        return True, "Now playing: " + song_name
    except Exception as e:
        return False, f"Error playing song: {str(e)}" 