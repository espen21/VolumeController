from ctypes import cast,POINTER
from comtypes import CLSCTX_ALL
import time
from pycaw.pycaw import AudioUtilities,IAudioEndpointVolume
import keyboard

devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_,CLSCTX_ALL,None)
volume = cast(interface,POINTER(IAudioEndpointVolume))





from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume

# Find the microphone device
mic_devices = AudioUtilities.GetMicrophone()
mic_interface = mic_devices.Activate(IAudioEndpointVolume._iid_,CLSCTX_ALL,None)
mic_volume = cast(mic_interface,POINTER(IAudioEndpointVolume))



# Mute the microphone ,1 = MUTE , 0 = Unmute mic
def toggle_mic(is_muted):
    if(is_muted):
        print("UnMuted")
        mic_volume.SetMute(0, None)
    else:
        print("Muted")
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
CURRENT_VOLUME = volume.GetMasterVolumeLevel()
print("Key binds\nUn/mute MIC:[Alt Gr]+[p]\nInCREASE VOLUME: [Alt Gr]+[i]\nDECREASE VOLUME: [Alt Gr]+[o]")
while(True):
    # Mic
    time.sleep(0.01)
    if(keyboard.is_pressed("Alt Gr+p")):
        is_muted = toggle_mic(is_muted)
        time.sleep(1)
    
    #InCREASE VOLUME
    if(keyboard.is_pressed("Alt Gr+i")):
        if CURRENT_VOLUME < MAX_VOL:
            time.sleep(0.01)
            CURRENT_VOLUME += 1
            volume.SetMasterVolumeLevel(CURRENT_VOLUME,None)
     #DECREASE VOLUME
    if(keyboard.is_pressed("Alt Gr+o")):
        if CURRENT_VOLUME >MIN_VOL:
            time.sleep(0.01)
            CURRENT_VOLUME -= 1
            volume.SetMasterVolumeLevel(CURRENT_VOLUME,None)
    