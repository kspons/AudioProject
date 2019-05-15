from pyfirmata2 import Arduino
import time
from ctypes import POINTER, cast
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, ISimpleAudioVolume
import time
import PySimpleGUIQt as sg  

class soundControl:

    def __init__(self):
        # Establish connection with arduino micro
        print("Connecting to arduino...")
        self.board = Arduino(Arduino.AUTODETECT)
        if self.board is None:
            print("Couldn't find arduino")
        else:
            print("Connected to arduino. Initializing volume control")
        
        # Setup master volume control
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        self.session_num = [None] * 4
        self.session_vol_prev = [0,0,0,0]
        self.session_vol = [0,0,0,0]
        self.session_num[0] = cast(interface, POINTER(IAudioEndpointVolume))

        self.board.analog[0].register_callback(self.app_vol_func_gen(0))
        self.board.samplingOn(80)
        self.board.analog[0].enable_reporting()

    def selectApps(self, appList):
        i = 0
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions[0:-1]: # last one is master volume & it has no name
            if session.Process.name() in appList: 
                i += 1
                self.session_num[i] = session._ctl.QueryInterface(ISimpleAudioVolume)
                print("found session for ", session.Process.name())
                self.board.analog[i].register_callback(self.app_vol_func_gen(i))
                self.board.analog[i].enable_reporting()
        while i < 5:
            i += 1
            self.board.analog[i].disable_reporting()

    def app_vol_func_gen(self, num):
        def appVol(data): # Master Volume
            if self.session_vol_prev[num] != data:
                if(num == 0):
                    self.session_num[num].SetMasterVolumeLevelScalar(data, None)
                else:
                    self.session_num[num].SetMasterVolume(data, None)
                self.session_vol_prev[num] = data
        return appVol
    
    def stop(self):
        self.board.samplingOff()
        self.board.exit()

def main():

    menu_def = ['BLANK', ['&Refresh', '---', 'E&xit']]  

    tray = sg.SystemTray(menu=menu_def, filename="orange.png")  
    exe_list = ["chrome.exe", "Spotify.exe", "Firefox.exe"]
    soundController = soundControl()
    soundController.selectApps(exe_list)

    try:
        while True:
            menu_item = tray.Read()
            print(menu_item)
            if menu_item == 'Exit':  
                soundController.stop()
                print("Exiting")
                break
            elif menu_item == 'Refresh':
                soundController.stop()
                soundController.__init__()
                soundController.selectApps(exe_list)
            continue

    except KeyboardInterrupt:
        soundController.stop()
        print("Exiting")


if __name__ == "__main__":
    main()