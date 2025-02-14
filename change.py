import sys
import subprocess
import os

name = "mi_voice_switch"
thispath = os.path.join(os.environ["LOCALAPPDATA"],"MG_mikan",name)

# インストールしたいパッケージのリスト（パッケージ名: バージョン）
packages = {
    "pycaw": "20240210",
    "keyboard": "0.13.5",
    "comtypes": "1.4.9",
    "win11toast": "0.35"
}
all_packages = [f"{pkg}=={ver}" for pkg, ver in packages.items()]

# 必要なパッケージのみインストール
if all_packages:
    print(f"Installing the following packages: {', '.join(all_packages)}")
    subprocess.run([sys.executable, "-m", "pip", "install", *all_packages], check=True)

import threading
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import keyboard
import win11toast

# オーディオデバイスの取得
devices = AudioUtilities.GetMicrophone()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))

# フラグ（ミュート切り替え用）

volume.SetMasterVolumeLevelScalar(1.0, None)  # 解除
is_muted = False



def show_notification(title, message,image):
    win11toast.toast(title, message, icon = image)

# キーボード入力時の処理
def on_press(key):
    global is_muted
    try:
        # if key.name == "caps lock" and 
        if keyboard.is_pressed("alt+m"):  # "caps lock" キーが押されたとき
            if is_muted:
                volume.SetMasterVolumeLevelScalar(1.0, None)  # 100% に戻す
                threading.Thread(target=show_notification,args=("みゅーとくん","マイクミュートを解除しました",os.path.join(thispath,"icon.png")),daemon=True).start()
            else:
                volume.SetMasterVolumeLevelScalar(0.0, None)  # ミュート
                threading.Thread(target=show_notification,args=("みゅーとくん","マイクをミュート状態にしました",os.path.join(thispath,"icon.png")),daemon=True).start()
            is_muted = not is_muted
    except AttributeError:
        pass    

# リスナーを開始
keyboard.hook(on_press)

keyboard.wait()