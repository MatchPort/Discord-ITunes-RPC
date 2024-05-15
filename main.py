import time
from pypresence import Presence
import win32com.client
import os

# Discord client ID from the developer portal
client_id = '1240303183989833779'
RPC = Presence(client_id)
RPC.connect()

def get_itunes_info():
    itunes = win32com.client.Dispatch('iTunes.Application')
    current_track = itunes.CurrentTrack
    if current_track:
        song_info = {
            'name': current_track.Name,
            'artist': current_track.Artist,
            'album': current_track.Album,
            'duration': current_track.Duration,
            'position': itunes.PlayerPosition,
            'artwork': None
        }
        # Try to get the artwork
        try:
            if current_track.Artwork.Count > 0:
                artwork = current_track.Artwork.Item(1)
                artwork_path = os.path.join(os.getcwd(), 'artwork.jpg')
                # Ensure the artwork file can be written to
                if os.path.exists(artwork_path):
                    os.remove(artwork_path)
                artwork.SaveArtworkToFile(artwork_path)
                song_info['artwork'] = artwork_path
        except Exception as e:
            print(f"Error saving artwork: {e}")
            song_info['artwork'] = None
        return song_info
    return None

while True:
    song_info = get_itunes_info()
    if song_info:
        # Use the asset key for the artwork
        large_image_key = 'artwork' if song_info['artwork'] else None
        
        RPC.update(
            state=f"by {song_info['artist']}",
            details=song_info['name'],
            large_image=large_image_key,  # Use the artwork asset key
            large_text=song_info['album'],
            start=time.time() - song_info['position']
        )
    else:
        RPC.clear()
    time.sleep(5)  # Update every 5 seconds
