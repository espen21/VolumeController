from ctypes import cast,POINTER
from comtypes import CLSCTX_ALL
import time
from pycaw.pycaw import AudioUtilities,IAudioEndpointVolume
import keyboard
import wave
import pyaudio

devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_,CLSCTX_ALL,None)
volume = cast(interface,POINTER(IAudioEndpointVolume))





from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume

# Find the microphone device
mic_devices = AudioUtilities.GetMicrophone()
mic_interface = mic_devices.Activate(IAudioEndpointVolume._iid_,CLSCTX_ALL,None)
mic_volume = cast(mic_interface,POINTER(IAudioEndpointVolume))






# Open the WAV file
def play_muted():
    with wave.open("wavs\mic_muted.wav", "rb") as wav_file:
        # Initialize the PyAudio object
        py_audio = pyaudio.PyAudio()

        # Open the audio stream
        stream = py_audio.open(format=py_audio.get_format_from_width(wav_file.getsampwidth()),
                                channels=wav_file.getnchannels(),
                                rate=wav_file.getframerate(),
                                output=True)

        # Play the audio data
        data = wav_file.readframes(1024)
        while data:
            stream.write(data)
            data = wav_file.readframes(1024)

        # Close the audio stream and PyAudio object
        stream.stop_stream()
        stream.close()
        py_audio.terminate()

def play_unmuted():
    with wave.open("wavs\mic_activated.wav", "rb") as wav_file:
    # Initialize the PyAudio object
        py_audio = pyaudio.PyAudio()

        # Open the audio stream
        stream = py_audio.open(format=py_audio.get_format_from_width(wav_file.getsampwidth()),
                                channels=wav_file.getnchannels(),
                                rate=wav_file.getframerate(),
                                output=True)

        # Play the audio data
        data = wav_file.readframes(1024)
        while data:
            stream.write(data)
            data = wav_file.readframes(1024)

        # Close the audio stream and PyAudio object
        stream.stop_stream()
        stream.close()
        py_audio.terminate()

# Mute the microphone ,1 = MUTE , 0 = Unmute mic
def toggle_mic(is_muted):
    if(is_muted):
        print("UnMuted")
        play_unmuted()
        mic_volume.SetMute(0, None)
        
    else:
        print("Muted")
        play_muted()
        mic_volume.SetMute(1, None)
      
    return  not is_muted

def set_IS_muted():
    value =mic_volume.GetMute()
    if value ==0:
        return False
    return True



is_muted = set_IS_muted()
MAX_VOL = 0 # högsta
MIN_VOL = -65 #
CURRENT_VOLUME = int(volume.GetMasterVolumeLevel())
print("Key binds\nUn/mute MIC:[Alt Gr]+[å]\nInCREASE VOLUME: [Alt Gr]+[ö]\nDECREASE VOLUME: [Alt Gr]+[ä]")
while(True):
    # Mic
    time.sleep(0.01)
    if(keyboard.is_pressed("Alt Gr+å")):
        try:
            is_muted = toggle_mic(is_muted)
            time.sleep(1)
        except Exception as e:
            print(e)
    #InCREASE VOLUME
    if(keyboard.is_pressed("Alt Gr+ö")):
        if CURRENT_VOLUME < MAX_VOL:
            try:
                time.sleep(0.01)
                CURRENT_VOLUME += 1
                volume.SetMasterVolumeLevel(CURRENT_VOLUME,None)
            except Exception as e:
                print(e)
     #DECREASE VOLUME
    if(keyboard.is_pressed("Alt Gr+ä")):
        if CURRENT_VOLUME >MIN_VOL:
            try:
                time.sleep(0.01)
                CURRENT_VOLUME -= 1
                volume.SetMasterVolumeLevel(CURRENT_VOLUME,None)
            except Exception as e:
                print(e)
