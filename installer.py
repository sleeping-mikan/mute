import sys
import subprocess
# インストールしたいパッケージのリスト（パッケージ名: バージョン）
packages = {
    "pillow": "10.3.0",
    "requests": "2.32.2"
}
all_packages = [f"{pkg}=={ver}" for pkg, ver in packages.items()]

# 必要なパッケージのみインストール
if all_packages:
    print(f"Installing the following packages: {', '.join(all_packages)}")
    subprocess.run([sys.executable, "-m", "pip", "install", *all_packages], check=True)

from tkinter import *
from tkinter import ttk
from PIL import Image,ImageTk
import os
from threading import Thread
import requests
import time
import zipfile


directory = os.path.join(os.environ["LOCALAPPDATA"],"MG_mikan","utils","assets","picture","icon")
icon_path = os.path.join(os.environ["LOCALAPPDATA"],"MG_mikan","utils","assets","picture","icon","haguruma.ico")
logo_path = os.path.join(os.environ["LOCALAPPDATA"],"MG_mikan","utils","assets","picture","icon","logo.png")
background_color = "#666666"

class _Progress():
    def __init__(self) -> None:
        self.progress_false = Image.new("RGBA",(1,8),(64,64,64,255))
        self.progress_true = Image.new("RGBA",(1,6),(255,195,0,255))
        self.progress_true_buffer = 1
        self.letter_y = 110
        self.offsetper_x = -400
        self.letter_offset_x = 0
        self.bar_size = 512
        self.side_image_size = 256
        self.now_progress = 0
    def set_property(self,letter: str,per: float) -> None:
        self.set_text(letter = letter, per = per)
        self.set_bar(per = per)
    def set_bar(self,per: float) -> None:
        """
        進捗バーを進める

        per     : 進捗度(per)
        """
        # percentから512段階に直す
        draw_area = per * 5.12 if per * 5.12 <= 512 else 512
        # すでに塗り終えたところから目標地点まで塗る
        for draw in range(self.now_progress,int(draw_area + 1)):
            self.canvas.create_image(self.side_image_size + 14 + draw + 1,self.side_image_size / 2 + self.progress_true_buffer,image = self.progress_true,anchor = NW)
        self.now_progress = draw
    def set_text(self,letter: str,delete: bool = True, per : float = 0.0) -> None:
        """
        進捗バーの上部に状態を表示する

        leter   : 表示する文字
        delete  : 過去の文字の削除(存在しない場合はerr)
        """
        pertxt = f"({per:6.2f}%)"
        # text を30文字に調整（超えたら省略、足りなければスペース埋め）
        if len(letter) > 30:
            formatted_text = letter[:29] + "…"  # 最後の1文字を「…」にする
        else:
            formatted_text = letter.rjust(30)  # 30文字未満ならスペース埋め
        if delete:
            self.canvas.delete(self.put_letter1)
            self.canvas.delete(self.put_letter2)
        self.put_letter1 = self.canvas.create_text(self.side_image_size + self.bar_size - self.letter_offset_x,self.letter_y,text = formatted_text,anchor = SE, fill="#FFFFFF")
        self.put_letter2 = self.canvas.create_text(self.side_image_size + self.bar_size - self.letter_offset_x + self.offsetper_x,self.letter_y,text = pertxt,anchor = SE, fill="#FFFFFF")
    def create(self,title: str,icon: str,image: str,window: Tk) -> None:
        """
        title   : ウィンドウの名前
        icon    : ウィンドウのアイコンパス
        image   : ウィンドウの左側に表示する画像のパス
        window  : TkまたはToplevelのインスタンス 但し、初期状態で呼び出す
        """
        def set_img(img) -> None:
            # __init__で読み込んだ画像をtkinter上に乗せるために変換
            self.progress_false = ImageTk.PhotoImage(image = self.progress_false)
            self.progress_true = ImageTk.PhotoImage(image = self.progress_true)
            # imgを読み込み、貼り付ける
            self.image = Image.open(fp=img)
            self.image = ImageTk.PhotoImage(image = self.image)
            self.canvas.create_image(0,0,image = self.image,anchor = NW)
            #初期状態の進捗バーを作成(灰色の横512 + 2px)
            for bar_draw_place in range(self.bar_size + 2):
                self.canvas.create_image(self.side_image_size + 14 + bar_draw_place,self.side_image_size / 2,image = self.progress_false,anchor = NW)
        self.root = window
        self.root.iconbitmap(icon)
        self.canvas = Canvas(self.root,width=self.side_image_size + 14 + self.bar_size + 2 + 5,height=self.side_image_size,background=background_color)
        self.canvas.pack()
        set_img(image)
        self.set_text(letter = "",delete = False)


class _Setting:
    def __init__(self, title, appname, downloads, entry, appico : str | None = None,startup: bool = False) -> None:
        self.title = title
        self.appname = appname
        self.downloads = downloads
        self.entry = entry
        self.appico = appico
        self.isstartup = startup
    def install(self):
        dumps_dir = os.path.expandvars(rf'%LOCALAPPDATA%\MG_mikan\{self.appname}')
        def create_short_cut(frm: os.path,to: os.path,name: os.path,icon: os.path):
            import comtypes.client
            #リンク先のファイル名
            target_file=frm
            #ショートカットを作成するパス
            save_path=to
            #WSHを生成
            wsh=comtypes.client.CreateObject("wScript.Shell",dynamic=True)
            #ショートカットの作成先を指定して、ショートカットファイルを開く。作成先のファイルが存在しない場合は、自動作成される。
            short=wsh.CreateShortcut(save_path)
            #以下、ショートカットにリンク先やコメントといった情報を指定する。
            #リンク先を指定
            short.TargetPath=target_file
            #コメントを指定する
            short.Description=name
            #引数を指定したい場合は、下記のコメントを解除して、引数を指定する。
            #short.arguments="/param1"
            #アイコンを指定したい場合は、下記のコメントを解除してアイコンのパスを指定する。
            short.IconLocation=icon
            #作業ディレクトリを指定したい場合は、下記のコメントを解除してディレクトリのパスを指定する。
            #short.workingDirectory="c:\\test\\"
            #ショートカットファイルを作成する
            short.Save()
        def _downloads(item: {str, os.path, bool, str}, per):
            """
            param -> item.url ダウンロードするurl
            param -> item.savedirectory ダウンロード先(アプリケーションパスlocalappdata/MG_mikan/<appname>/が起点)
            param -> item.iszip ダウンロードファイルを展開する必要があるか
            param -> item.viewtxt: <ダウンロード中に表示する文字> 
            """
            # テキストを表示
            self.progress.set_property(letter = item["viewtxt"], per = per)
            # ダウンロード
            count = 10
            for i in range(1,count + 1):
                write_path = os.path.join(dumps_dir, item["savedirectory"])
                os.makedirs(os.path.dirname(write_path), exist_ok=True)
                url = item["url"]  # ここに適切なURLを設定
                response = requests.get(url, stream=True)
                if response.status_code == 200:
                    url_data = requests.get(url).content
                    with open(write_path ,mode='wb') as f: # wb でバイト型を書き込める
                        f.write(url_data)
                    # ZIPファイルの場合、解凍
                    if item["iszip"]:
                        extract_path = os.path.splitext(write_path)[0]  # .zip を除いたフォルダを作成
                        os.makedirs(extract_path, exist_ok=True)
                        
                        with zipfile.ZipFile(write_path, 'r') as zip_ref:
                            zip_ref.extractall(extract_path)
                        
                        # ZIPファイルを削除する場合
                        os.remove(write_path)
                        
                        print(f"Extracted to: {extract_path}")
                    return
                else:
                    time.sleep(0.25)
                    print(f"ダウンロードに失敗しました。残り{count - i}回再試行します。")
            print(f"インストールに失敗しました。時間をおいて再度実行してください。時間をおいても解決しない場合、制作者に問合せてください。")
            
        # rootが存在しなければ作成しておく
        if not os.path.exists(dumps_dir):
            os.makedirs(dumps_dir)

        # 要求されたものを全てインストール
        for item in range(len(self.downloads)):
            _downloads(self.downloads[item], item / len(self.downloads) * 100)

        if self.appico == None:
            ico = icon_path
        else:
            ico = os.path.join(dumps_dir,"appico.ico")
            self.progress.set_property(letter = "アイコンをダウンロード中...",per=100.0)
            url = self.appico # ここに適切なURLを設定
            response = requests.get(url, stream=True)
            if response.status_code == 200:
                url_data = requests.get(url).content
                with open(os.path.join(dumps_dir,"appico.ico") ,mode='wb') as f: # wb でバイト型を書き込める
                    f.write(url_data)
            else:
                print(f"インストールに失敗しました。時間をおいて再度実行してください。時間をおいても解決しない場合、制作者に問合せてください。")
                return

        if self.button_data[0]:
            self.progress.set_property(letter = "ショートカットを生成中...",per=100.0)
            create_short_cut(os.path.expanduser(f"%LOCALAPPDATA%\\MG_mikan\\{self.appname}\\{self.entry}"),os.path.expanduser(f"~\\Desktop\\{self.appname}.lnk"),f"{self.appname}",ico)
        if self.isstartup:
            self.progress.set_property(letter = "スタートアップに登録中",per=100.0)
            create_short_cut(f"{dumps_dir}/{self.entry}",os.getenv('APPDATA') + rf"\Microsoft\Windows\Start Menu\Programs\Startup\{self.appname}.lnk",f"{self.appname}",ico)

        if self.button_data[1]:
            self.progress.set_property(letter = "アプリケーションメニューに登録中",per=100.0)
            create_short_cut(f"{dumps_dir}/{self.entry}",os.getenv('APPDATA') + rf"\Microsoft\Windows\Start Menu\Programs\{self.appname}.lnk",f"{self.appname}",ico)

        self.progress.set_property(letter = "ダウンロードが完了しました！",per=100.0)
    def start_process(self):
        self.button_data = [self.short_bln.get(),self.menu_bln.get()]
        self.canvas.destroy()
        self.progress.create(f"{self.appname} SetUp",icon=icon_path,image=logo_path,window=self.root)
        Thread(target = self.install,daemon = True).start()
    def _set_obj(self):
        # %LOCALAPPDATA%\haguruma
        self.image = Image.open(fp=os.path.join(os.environ["LOCALAPPDATA"],"MG_mikan","utils","assets","picture","icon","logo.png"))
        self.image = ImageTk.PhotoImage(image = self.image)
        self.canvas.create_image(0,0,image = self.image,anchor = NW)
        # チェックONにする
        self.short_bln = BooleanVar()
        self.short_bln.set(True)
        self.menu_bln = BooleanVar()
        self.menu_bln.set(True)
        style = ttk.Style()
        style.configure("Custom.TCheckbutton", background=background_color, foreground = "#FFFFFF")
        self.short_cut_button = ttk.Checkbutton(self.root,text="デスクトップにショートカットを作成する",variable=self.short_bln,style="Custom.TCheckbutton")
        self.short_cut_button.place(x = 200,y = 180)
        self.menu_cut_button = ttk.Checkbutton(self.root,text="アプリケーションメニューに表示する",variable=self.menu_bln,style="Custom.TCheckbutton")
        self.menu_cut_button.place(x = 200,y = 200)
        self.start_button = ttk.Button(self.root,text = "インストール",command=self.start_process)
        self.start_button.place(x = 315,y = 225)
    def create(self):
        self.root = Tk()
        self.root.title(self.title)
        self.progress = _Progress()
        self.root.iconbitmap(os.path.join(os.environ["LOCALAPPDATA"],"MG_mikan","utils","assets","picture","icon","haguruma.ico"))
        self.canvas = Canvas(self.root,width=400,height=256,background=background_color)
        self.canvas.pack()
        self._set_obj()
        self.root.mainloop()

class Installer():
    def __init__(self, title, appname, downloads, entry, appico: str | None = None, startup: bool = False) -> None:
        """
        param -> title タイトルバーに表示する名前
        param -> appname インストールするソフトウェアの名前
        param -> downloads [{url: <str>, savedirectory: <os.path>, iszip: <bool>, viewtxt: <str>}]
            param -> downloads.url ダウンロードするurl
            param -> downloads.savedirectory ダウンロード先(アプリケーションパスlocalappdata/MG_mikan/<appname>/が起点)
            param -> downloads.iszip ダウンロードファイルを展開する必要があるか
            param -> downloads.viewtxt: <ダウンロード中に表示する文字> 
        param -> entry エントリポイントが存在するプログラム/実行ファイル
        param -> appico アプリアイコンurl
        param -> startup スタートアップに登録するか
        """
        if not(isinstance(title, str) 
                and isinstance(appname, str) 
                and isinstance(downloads, list)
                and all(
                    isinstance(item, dict)  # itemが辞書であることを確認
                    and "url" in item and isinstance(item["url"], str)
                    and "savedirectory" in item and isinstance(item["savedirectory"], str)
                    and "iszip" in item and isinstance(item["iszip"], bool)
                    and "viewtxt" in item and isinstance(item["viewtxt"], str)
                    for item in downloads
                ) 
                and isinstance(entry, str) 
                and isinstance(appico, (str, type(None))) 
                and isinstance(startup, bool)
            ):
            raise ValueError
        # ファイルが存在しなければ、ダウンロードする
        self._download()
        self.progress = _Progress()
        self.setting = _Setting(title, appname, downloads,entry,appico,startup)
        try:
            self.setting.create()
        except:
            print("アプリケーションが終了されました")
            exit()
    def _download(self):
        if not os.path.exists(directory):
            # ディレクトリを作成
            os.makedirs(directory)
        # icoダウンロード
        if not os.path.exists(icon_path):
            import requests
            url = "https://www.dropbox.com/scl/fi/0a2qker16nlyz8xb3io63/mic.ico?rlkey=mlioybuvywk6i1tbkcnt0dd9p&st=fcr01e64&dl=1"  # ここに適切なURLを設定
            response = requests.get(url, stream=True)
            if response.status_code == 200:
                url_data = requests.get(url).content
                with open(icon_path ,mode='wb') as f: # wb でバイト型を書き込める
                    f.write(url_data)
            else:
                print("ファイルをダウンロードできません")
        # logoダウンロード
        if not os.path.exists(logo_path):
            import requests
            url = "https://www.dropbox.com/scl/fi/qaac5n8gvv25zfo11hcvj/group_gradation_.png?rlkey=qtu3pctg1ds1ogha0u9kskgxq&st=z3vn4v39&dl=1"  # ここに適切なURLを設定
            response = requests.get(url, stream=True)
            if response.status_code == 200:
                url_data = requests.get(url).content
                with open(logo_path ,mode='wb') as f: # wb でバイト型を書き込める
                    f.write(url_data)
            else:
                print("アイコンのダウンロードに失敗しました。")


if __name__ == "__main__":
    items = [
        {
            "url": "https://github.com/sleeping-mikan/mute/archive/refs/heads/main.zip",
            "savedirectory": "main.zip",
            "iszip": True,
            "viewtxt": "repos.zipをダウンロード中"
        }
    ]
    installer = Installer(
            title = "Voice Changer Installer",
            appname="mi_voice_switch",
            downloads=items, 
            entry = "main.pyw", 
            appico="https://www.dropbox.com/scl/fi/cw75lmhpkala097rkyzzf/haguruma.ico?rlkey=eaw6236xrpl1dhnwrmw0dvayc&st=76eknxvh&dl=1"
        )
    subprocess.run(["py", os.environ["LOCALAPPDATA"],"MG_mikan","mi_voice_switch","main.pyw"])