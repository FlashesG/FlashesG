import customtkinter as ctk
import subprocess
import threading
import json
import os
import sys
import webbrowser
import base64
import psutil
import time
from datetime import datetime
import tkinter as tk
from tkinter import scrolledtext
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from cryptography.fernet import Fernet
import functools

# アプリケーションのパス検出


def get_base_path():
    """アプリケーションの基本パスを取得する（EXE実行時と通常実行時で異なる）"""
    if getattr(sys, 'frozen', False):
        # EXE形式で実行された場合
        return os.path.dirname(sys.executable)
    else:
        # 通常のPythonスクリプトとして実行された場合
        return os.path.dirname(os.path.abspath(__file__))

# 設定ファイルのパスとログファイルのパス


def get_config_path():
    """設定ファイルのパスを取得"""
    return os.path.join(get_base_path(), 'mapcamera_config.json')


def get_log_path():
    """ログファイルのパスを取得"""
    return os.path.join(get_base_path(), 'mapcamera_log.txt')

# 鍵ファイルのパスを取得


def get_key_path():
    """暗号化キーのパスを取得"""
    base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, 'encryption.key')

# 鍵の生成または読み込み


def get_encryption_key():
    """暗号化キーを取得または生成"""
    key_path = get_key_path()

    # キーファイルが存在しない場合は新しく生成
    if not os.path.exists(key_path):
        key = Fernet.generate_key()
        with open(key_path, 'wb') as key_file:
            key_file.write(key)
        return key

    # 既存のキーファイルから読み込み
    with open(key_path, 'rb') as key_file:
        return key_file.read()


def encrypt_password(password):
    """パスワードを暗号化"""
    if not password:
        return ""
    try:
        key = get_encryption_key()
        cipher = Fernet(key)
        encrypted_data = cipher.encrypt(password.encode())
        return base64.urlsafe_b64encode(encrypted_data).decode()
    except Exception as e:
        print(f"暗号化エラー: {str(e)}")
        return ""


def decrypt_password(encrypted_data):
    """暗号化されたパスワードを復号化"""
    if not encrypted_data:
        return ""
    try:
        key = get_encryption_key()
        cipher = Fernet(key)
        decrypted_data = cipher.decrypt(
            base64.urlsafe_b64decode(encrypted_data.encode()))
        return decrypted_data.decode()
    except Exception as e:
        print(f"復号化エラー: {str(e)}")
        return ""


# エラーメッセージ定数
ERROR_MESSAGES = {
    "chrome_not_found": "Googleクロームがインストールされていないか、指定されたパスが間違っています。設定を確認してください。",
    "password_empty": "パスワードが設定されていません。パスワード欄にマップカメラサイトのパスワードを入力して「保存」ボタンをクリックしてください。",
    "tab_not_found": "マップカメラのタブが見つかりません。「マップカメラサイトを開く」ボタンをクリックしてください。",
    "connection_error": "インターネット接続に問題があるか、マップカメラのサイトにアクセスできません。接続を確認してから再試行してください。",
    "element_not_found": "ページ上の要素が見つかりませんでした。マップカメラのサイト構造が変更された可能性があります。",
    "recaptcha_error": "reCAPTCHA認証で問題が発生しました。手動でチェックを入れてください。",
    "max_retries": "最大試行回数に達しました。サイトの応答が遅いか、構造が変更されている可能性があります。",
}

# 状態表示の色定義
STATUS_COLORS = {
    "info": "black",
    "success": "green",
    "warning": "#FF8C00",  # ダークオレンジ
    "error": "#E57373"     # 明るい赤
}

# エラー処理デコレータ


def gui_error_handler(operation=None):
    """GUIメソッドのエラーを捕捉して処理するデコレータ"""
    def decorator(func):
        @functools.wraps(func)
        def wrapper(self, *args, **kwargs):
            try:
                return func(self, *args, **kwargs)
            except Exception as e:
                op_name = operation or func.__name__
                error_msg = f"{op_name}でエラーが発生: {str(e)} ({type(e).__name__})"
                self.log(error_msg)
                # 詳細ログが有効な場合はスタックトレースも表示
                if hasattr(self, 'verbose_var') and self.verbose_var.get():
                    import traceback
                    trace_info = traceback.format_exc()
                    self.log(f"詳細なスタックトレース:\n{trace_info}")
                # UIに通知
                self.update_status(f"エラーが発生しました: {str(e)}", "error")
                return None
        return wrapper
    return decorator

# ロガークラス - ログ機能を一元管理


class Logger:
    """ログ機能の管理クラス"""

    def __init__(self, log_widget=None, log_file=None):
        self.log_widget = log_widget
        self.log_file = log_file

    def log(self, message, level="INFO"):
        """ログを記録して表示"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        formatted_message = f"[{timestamp}] [{level}] {message}"

        # コンソールに出力
        print(formatted_message)

        # GUIのログエリアに表示
        if self.log_widget:
            try:
                self.log_widget.config(state=tk.NORMAL)
                self.log_widget.insert(tk.END, formatted_message + "\n")
                self.log_widget.see(tk.END)  # 自動スクロール
                self.log_widget.config(state=tk.DISABLED)
            except Exception as e:
                print(f"ログウィジェット更新エラー: {str(e)}")

        # ファイルに記録
        if self.log_file:
            try:
                with open(self.log_file, 'a', encoding='utf-8') as f:
                    f.write(formatted_message + "\n")
            except Exception as e:
                print(f"ログファイル書き込みエラー: {str(e)}")

    def info(self, message):
        """情報メッセージをログに記録"""
        self.log(message, "INFO")

    def warning(self, message):
        """警告メッセージをログに記録"""
        self.log(message, "WARNING")

    def error(self, message):
        """エラーメッセージをログに記録"""
        self.log(message, "ERROR")

    def success(self, message):
        """成功メッセージをログに記録"""
        self.log(message, "SUCCESS")

# 設定管理クラス - 設定の読み込みと保存を担当


class ConfigManager:
    """設定の読み込みと保存を管理"""

    def __init__(self, config_path, logger=None):
        self.config_path = config_path
        self.logger = logger
        self.default_config = {
            'password': '',
            'chrome_path': r'C:\Program Files\Google\Chrome\Application\chrome.exe',
            'profile_directory': 'Profile 7',
            'debug_port': '9222',
            'verbose_log': False,
            'wait_time': 5,
            'max_retries': 5,
            'payment_method': 'daibiki',
            'poll_frequency': 0.2,
            'page_load_timeout': 20,
            'script_timeout': 15,
            'force_stop_timeout': 2000,
            'auto_switch_tab': False  # 購入完了後のタブ自動切り替え（現在は無効）
        }
        self.config = self.load()

    def log(self, message):
        """ログメッセージを記録"""
        if callable(self.logger):
            self.logger(message)

    def load(self):
        """設定ファイルを読み込む"""
        if not os.path.exists(self.config_path):
            self.log("設定ファイルがありません。デフォルト設定を使用します。")
            return self.default_config.copy()

        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
                # デフォルト値とマージ
                result = self.default_config.copy()
                result.update(config)

                # パスワード復号化
                if result.get('password'):
                    result['password'] = decrypt_password(result['password'])

                self.log("設定ファイルを読み込みました")
                return result
        except Exception as e:
            self.log(f"設定ファイルの読み込みに失敗しました: {str(e)}")
            return self.default_config.copy()

    def save(self, config=None):
        """設定ファイルを保存する"""
        config = config or self.config

        # 保存用にコピーを作成
        save_config = config.copy()

        # パスワード暗号化
        if save_config.get('password'):
            save_config['password'] = encrypt_password(save_config['password'])

        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(save_config, f, ensure_ascii=False, indent=2)
            self.log("設定を保存しました")
            return True
        except Exception as e:
            self.log(f"設定の保存に失敗しました: {str(e)}")
            return False

# プロセス管理クラス - Chromeドライバープロセスの管理


class ProcessManager:
    """システムプロセスの管理を担当"""

    def __init__(self, logger=None):
        self.logger = logger

    def log(self, message):
        """ログメッセージを記録"""
        if self.logger:
            self.logger.info(message)
        else:
            print(message)

    def cleanup_chrome_drivers(self):
        """古いChromeDriverプロセスをクリーンアップ"""
        try:
            self.log("古いWebDriverプロセスをチェックしています...")
            count = 0
            terminated_pids = []

            # 終了させるプロセスのリストを作成
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    # ChromeDriverプロセスを探す
                    if 'chromedriver' in proc.info['name'].lower():
                        # 終了リストに追加
                        terminated_pids.append(proc.info['pid'])
                        try:
                            proc.terminate()
                            self.log(
                                f"ChromeDriverプロセス(PID: {proc.info['pid']})の終了を要求しました")
                            count += 1
                        except psutil.AccessDenied:
                            self.log(
                                f"権限不足: PID {proc.info['pid']}の終了には管理者権限が必要かもしれません")
                        except Exception as e:
                            self.log(
                                f"プロセス終了エラー (PID: {proc.info['pid']}): {str(e)}")
                except (psutil.NoSuchProcess, psutil.ZombieProcess):
                    pass
                except Exception as e:
                    self.log(f"プロセス情報取得エラー: {str(e)}")

            # 少し待機してプロセスの終了を確認
            if terminated_pids:
                self.log(f"{len(terminated_pids)}個のプロセスの終了を要求しました。終了を確認中...")
                time.sleep(0.5)  # 0.5秒待機

                # 終了しなかったプロセスを強制終了
                for pid in terminated_pids:
                    try:
                        proc = psutil.Process(pid)
                        if proc.is_running():
                            self.log(f"PID {pid}が終了していません。強制終了を試みます...")
                            proc.kill()  # より強力な終了方法
                    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                        # NoSuchProcessは正常（すでに終了している）
                        pass
                    except Exception as e:
                        self.log(f"プロセス強制終了エラー (PID: {pid}): {str(e)}")

                # さらに待機して最終確認
                time.sleep(0.5)
                still_running = []
                for pid in terminated_pids:
                    try:
                        proc = psutil.Process(pid)
                        if proc.is_running():
                            still_running.append(pid)
                    except (psutil.NoSuchProcess, psutil.ZombieProcess):
                        # 正常に終了
                        pass
                    except Exception:
                        pass

                if still_running:
                    self.log(
                        f"警告: {len(still_running)}個のプロセスが終了しませんでした。PIDs: {still_running}")
                else:
                    self.log(f"すべてのプロセス({count}個)が正常に終了しました")
            else:
                self.log("クリーンアップが必要な古いプロセスはありませんでした")

        except Exception as e:
            self.log(f"プロセスクリーンアップエラー: {str(e)}")

    def check_chrome_running(self, debug_port):
        """指定のデバッグポートでChromeが実行中かどうかを確認"""
        try:
            for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
                try:
                    proc_info = proc.info
                    if 'chrome' in proc_info['name'].lower():
                        # コマンドラインでデバッグポートを確認
                        if proc_info['cmdline'] and any(f"--remote-debugging-port={debug_port}" in cmd for cmd in proc_info['cmdline']):
                            return True
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    pass
            return False
        except Exception as e:
            self.log(f"Chrome実行確認エラー: {str(e)}")
            return False

# 自動化タスク管理クラス - 自動化処理を別スレッドで実行


class AutomationTask:
    """バックグラウンドで実行する自動化タスク"""

    def __init__(self, logger=None, ui_updater=None, app=None):
        self.logger = logger
        self.ui_updater = ui_updater or (lambda callback, delay=0: None)
        self.running = False
        self.stop_requested = False
        self.thread = None
        self.automation = None
        self.app = app  # アプリケーション参照を追加

    def log(self, message):
        """ログメッセージを記録"""
        if self.logger:
            self.logger.info(message)
        else:
            print(message)

    def run(self, target, args=(), kwargs=None, on_complete=None):
        """タスクをバックグラウンドで実行"""
        if self.running:
            self.log("既にタスクが実行中です")
            return False

        # 状態を完全にリセット（追加）
        self.stop_requested = False
        if hasattr(self, 'automation') and self.automation:
            if hasattr(self.automation, 'stop_requested'):
                self.automation.stop_requested = False

        kwargs = kwargs or {}

        def wrapped_target():
            try:
                self.log("タスクを開始します")
                self.running = True

                # 対象関数を実行
                result = target(*args, **(kwargs or {}))

                # 完了処理
                if not self.stop_requested and callable(on_complete):
                    self.ui_updater(lambda: on_complete(result), 0)

                self.log("タスクが完了しました")
                return result
            except Exception as e:
                self.log(f"タスク実行中にエラーが発生: {str(e)}")
                return None
            finally:
                # 状態をリセットして UI 更新
                was_running = self.running  # 実行状態を記録
                self.running = False
                self.stop_requested = False

                # UI 状態リセット用のコールバックを追加
                if was_running:  # 実際に実行中だった場合のみ UI 更新
                    self.ui_updater(self._reset_ui_state, 100)  # 少し遅延させて確実に実行

        # スレッド作成と開始
        self.thread = threading.Thread(target=wrapped_target, daemon=True)
        self.thread.start()
        return True

    def stop(self):
        """タスクの停止をリクエスト"""
        if not self.running:
            return False

        self.log("タスクの停止をリクエストしています")
        self.stop_requested = True

        # 自動化クラスにも停止リクエストを伝達
        if self.automation and hasattr(self.automation, 'request_stop'):
            self.automation.request_stop()

        return True

    def force_stop(self):
        """タスクを強制的に停止（スレッド終了待機機能付き）"""
        if not self.running:
            return False

        self.log("タスクを強制停止します")
        self.stop_requested = True

        # 自動化クラスにも停止を伝達
        if self.automation:
            try:
                if hasattr(self.automation, 'stop_requested'):
                    self.automation.stop_requested = True
                if hasattr(self.automation, 'is_shutting_down'):
                    self.automation.is_shutting_down = True

                # スレッドの実行状態を確実に停止するためのシグナル設定
                if hasattr(self.automation, 'driver'):
                    self.automation.driver = None  # ドライバー参照を意図的に無効化

                # クリーンアップ処理
                if hasattr(self.automation, 'cleanup'):
                    self.automation.cleanup()
            except Exception as e:
                self.log(f"自動化クラスの停止処理でエラー: {str(e)}")

        # 実行状態をリセット
        self.running = False

        # スレッドが終了するまで少し待機（非ブロッキング）
        if self.thread and self.thread.is_alive():
            self.thread.join(0.5)  # 最大0.5秒待機

        return True

    def _reset_ui_state(self):
        """UI状態をリセットするコールバック"""
        # アプリケーションのUI状態リセット用イベント発火
        if hasattr(self, 'app') and self.app:
            # アプリケーションのUI更新メソッドを呼び出す
            self.app.reset_ui_state_after_task()

# UIユーティリティクラス - UI構築を支援


class UIFactory:
    """一貫したUI部品を作成するファクトリー"""

    def __init__(self, app):
        self.app = app

    def create_label(self, parent, text, **kwargs):
        """標準スタイルのラベルを作成"""
        font = kwargs.pop('font', self.app.normal_font)
        label = ctk.CTkLabel(parent, text=text, font=font, **kwargs)
        return label

    def create_button(self, parent, text, command, **kwargs):
        """標準スタイルのボタンを作成"""
        font = kwargs.pop('font', self.app.button_font)
        button = ctk.CTkButton(
            parent, text=text, command=command, font=font, **kwargs)
        return button

    def create_frame(self, parent, **kwargs):
        """標準スタイルのフレームを作成"""
        frame = ctk.CTkFrame(parent, **kwargs)
        return frame

    def create_entry(self, parent, **kwargs):
        """標準スタイルのテキスト入力を作成"""
        font = kwargs.pop('font', self.app.normal_font)
        entry = ctk.CTkEntry(parent, font=font, **kwargs)
        return entry

    def create_checkbox(self, parent, text, variable, **kwargs):
        """標準スタイルのチェックボックスを作成"""
        font = kwargs.pop('font', self.app.normal_font)
        checkbox = ctk.CTkCheckBox(
            parent, text=text, variable=variable, font=font, **kwargs)
        return checkbox

    def create_combobox(self, parent, values, **kwargs):
        """標準スタイルのコンボボックスを作成"""
        font = kwargs.pop('font', self.app.normal_font)
        combobox = ctk.CTkComboBox(parent, values=values, font=font, **kwargs)
        return combobox

# ダイアログ管理クラス - 各種ダイアログを管理


class DialogManager:
    """対話的なダイアログの管理"""

    def __init__(self, root, ui_factory):
        self.root = root
        self.ui = ui_factory

    def show_info(self, message):
        """情報ダイアログを表示"""
        dialog = ctk.CTkToplevel(self.root)
        dialog.title("情報")
        dialog.geometry("400x150")
        dialog.grab_set()

        self.ui.create_label(
            dialog,
            text=message
        ).pack(padx=20, pady=(20, 10), fill=tk.BOTH, expand=True)

        self.ui.create_button(
            dialog,
            text="OK",
            command=dialog.destroy
        ).pack(padx=20, pady=20)

        return dialog

    def show_error(self, message, error_key=None):
        """エラーダイアログを表示"""
        if error_key:
            message = ERROR_MESSAGES.get(error_key, message)

        dialog = ctk.CTkToplevel(self.root)
        dialog.title("エラー")
        dialog.geometry("450x200")
        dialog.grab_set()

        # エラーアイコン作成
        icon_size = 60
        icon_frame = ctk.CTkFrame(
            dialog, width=icon_size, height=icon_size, fg_color="transparent")
        icon_frame.pack(pady=(15, 5))

        canvas = tk.Canvas(icon_frame, width=icon_size, height=icon_size,
                           bg=dialog.cget('bg'), highlightthickness=0)
        canvas.pack()

        # 赤い円を描画
        canvas.create_oval(5, 5, icon_size-5, icon_size -
                           5, fill="#E53935", outline="")

        # 白い「×」マークを描画
        line_width = 4
        margin = 18
        canvas.create_line(margin, margin, icon_size-margin, icon_size-margin,
                           fill="white", width=line_width, capstyle=tk.ROUND)
        canvas.create_line(icon_size-margin, margin, margin, icon_size-margin,
                           fill="white", width=line_width, capstyle=tk.ROUND)

        # メッセージラベル
        self.ui.create_label(
            dialog,
            text=message,
            text_color="#E57373",
            wraplength=400
        ).pack(padx=20, pady=(10, 15), fill=tk.BOTH, expand=True)

        # OKボタン
        self.ui.create_button(
            dialog,
            text="OK",
            command=dialog.destroy,
            fg_color="#E57373",
            hover_color="#EF5350",
            width=100,
            height=35
        ).pack(padx=20, pady=15)

        # エスケープキーでダイアログを閉じる
        dialog.bind("<Escape>", lambda event: dialog.destroy())

        # ダイアログの中央にフォーカス
        dialog.after(100, lambda: dialog.focus_set())

        return dialog

    def show_confirm(self, message):
        """確認ダイアログを表示して結果を返す"""
        result = [False]  # リストを使って結果を格納

        dialog = ctk.CTkToplevel(self.root)
        dialog.title("確認")
        dialog.geometry("400x150")
        dialog.grab_set()

        self.ui.create_label(
            dialog,
            text=message
        ).pack(padx=20, pady=(20, 10), fill=tk.BOTH, expand=True)

        button_frame = ctk.CTkFrame(dialog)
        button_frame.pack(fill=tk.X, padx=20, pady=20)

        def on_yes():
            result[0] = True
            dialog.destroy()

        def on_no():
            result[0] = False
            dialog.destroy()

        self.ui.create_button(
            button_frame,
            text="はい",
            command=on_yes
        ).pack(side=tk.LEFT, padx=10)

        self.ui.create_button(
            button_frame,
            text="いいえ",
            command=on_no,
            fg_color="#E57373",
            hover_color="#EF5350"
        ).pack(side=tk.RIGHT, padx=10)

        # ダイアログが閉じられるまで待機
        self.root.wait_window(dialog)

        return result[0]

# メインアプリケーションクラス


class MapCameraGUI:
    def __init__(self):
        # テーマ設定
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        # メインウィンドウの作成
        self.root = ctk.CTk()
        self.root.title("マップカメラ自動購入ツール")
        self.root.geometry("800x700")
        self.root.minsize(700, 700)

        # フォント設定
        self.setup_fonts()

        # ユーティリティクラスの初期化
        self.ui = UIFactory(self)

        # ロガーの初期化（UIコンポーネント作成後に完全初期化）
        self.logger = Logger(log_file=get_log_path())

        # 設定マネージャーの初期化
        self.config_manager = ConfigManager(get_config_path(), self.logger)
        self.config = self.config_manager.config

        # プロセスマネージャーの初期化
        self.process_manager = ProcessManager(self.logger)

        # ダイアログマネージャーの初期化
        self.dialog = DialogManager(self.root, self.ui)

        # タスク管理の初期化
        self.task = AutomationTask(
            logger=self.logger,
            ui_updater=self.ui_update_wrapper,
            app=self  # 自分自身への参照を渡す
        )

        # 変数初期化
        self.chrome_running = False
        self.automation = None
        self.force_stop_timer = None
        self.verbose_var = tk.BooleanVar(
            value=self.config.get('verbose_log', False))
        # 監視状態の初期化
        self.is_monitoring = False

        # UIの作成
        self.create_main_layout()

        # ロガーのUIコンポーネント登録（UIコンポーネント作成後）
        self.logger.log_widget = self.log_text

        # 初期ログ
        self.log("マップカメラ自動購入ツールを起動しました")
        self.update_status("準備完了しました。「Chromeを起動」ボタンをクリックしてください。", "info")

        # ツールチップの追加
        self.add_tooltips()

        # Chromeの状態確認タイマーは起動時には開始しない

    def setup_fonts(self):
        """フォントの設定"""
        try:
            self.title_font = ctk.CTkFont(
                family="游ゴシック", size=24, weight="bold")
            self.header_font = ctk.CTkFont(
                family="メイリオ", size=18, weight="bold")
            self.normal_font = ctk.CTkFont(family="メイリオ", size=12)
            self.button_font = ctk.CTkFont(family="メイリオ", size=13)
        except Exception as e:
            print(f"フォント初期化エラー: {str(e)}")
            # フォールバックとしてデフォルトフォントを使用
            self.title_font = ctk.CTkFont(size=24, weight="bold")
            self.header_font = ctk.CTkFont(size=18, weight="bold")
            self.normal_font = ctk.CTkFont(size=12)
            self.button_font = ctk.CTkFont(size=13)

    def create_main_layout(self):
        """メインレイアウトの作成"""
        # メインフレーム
        self.main_frame = self.ui.create_frame(self.root)
        self.main_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # ボタンフレーム
        self.button_frame = self.ui.create_frame(self.root)
        self.button_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

        # グリッドの重みを設定
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        # タイトルラベル
        self.ui.create_label(
            self.main_frame,
            text="マップカメラ自動購入ツール",
            font=self.title_font
        ).pack(pady=10)

        # 各セクションの作成
        self.create_browser_section()
        self.create_operation_section()
        self.create_status_section()
        self.create_settings_section()
        self.create_log_section()

        # 終了ボタン
        self.exit_button = self.ui.create_button(
            self.button_frame,
            text="終了",
            command=self.exit_app,
            fg_color="#E57373",
            hover_color="#EF5350"
        )
        self.exit_button.pack(fill=tk.X, pady=5, padx=10)

    def create_browser_section(self):
        """ブラウザ設定セクションの作成"""
        browser_frame = self.ui.create_frame(self.main_frame)
        browser_frame.pack(fill=tk.X, padx=10, pady=5)

        self.ui.create_label(
            browser_frame,
            text="ブラウザ設定",
            font=self.header_font
        ).pack(anchor="w", padx=10, pady=5)

        browser_controls = self.ui.create_frame(browser_frame)
        browser_controls.pack(fill=tk.X, padx=10, pady=5)

        self.chrome_button = self.ui.create_button(
            browser_controls,
            text="Chromeを起動",
            command=self.start_chrome
        )
        self.chrome_button.pack(side=tk.LEFT, padx=10, pady=5)

        self.ui.create_label(
            browser_controls,
            text="ステータス:"
        ).pack(side=tk.LEFT, padx=10, pady=5)

        self.chrome_status = self.ui.create_label(
            browser_controls,
            text="停止中",
            text_color="red"
        )
        self.chrome_status.pack(side=tk.LEFT, padx=5, pady=5)

    def create_operation_section(self):
        """操作セクションの作成"""
        operation_frame = self.ui.create_frame(self.main_frame)
        operation_frame.pack(fill=tk.X, padx=10, pady=5)

        self.ui.create_label(
            operation_frame,
            text="基本操作",
            font=self.header_font
        ).pack(anchor="w", padx=10, pady=5)

        operation_controls = self.ui.create_frame(operation_frame)
        operation_controls.pack(fill=tk.X, padx=10, pady=5)

        # 商品一覧から購入開始ボタン
        self.start_list_button = self.ui.create_button(
            operation_controls,
            text="商品一覧から購入開始",
            command=self.start_from_product_list,
            state="disabled"
        )
        self.start_list_button.grid(
            row=0, column=0, padx=10, pady=5, sticky="ew")

        # 商品ページから購入開始ボタン
        self.start_single_button = self.ui.create_button(
            operation_controls,
            text="商品ページから購入開始",
            command=self.start_from_product_page,
            state="disabled"
        )
        self.start_single_button.grid(
            row=0, column=1, padx=10, pady=5, sticky="ew")

        # マップカメラサイトを開くボタン
        self.open_site_button = self.ui.create_button(
            operation_controls,
            text="マップカメラサイトを開く",
            command=self.open_mapcamera_site
        )
        self.open_site_button.grid(
            row=1, column=0, columnspan=2, padx=10, pady=5, sticky="ew")

        # グリッドの列の重み設定
        operation_controls.grid_columnconfigure(0, weight=1)
        operation_controls.grid_columnconfigure(1, weight=1)

        # 停止ボタン
        self.stop_button = self.ui.create_button(
            operation_controls,
            text="処理を停止",
            command=self.stop_automation,
            state="disabled",
            fg_color="#E57373",
            hover_color="#EF5350"
        )
        self.stop_button.grid(row=2, column=0, columnspan=2,
                              padx=10, pady=5, sticky="ew")

        # 連続購入モードボタン
        self.continuous_mode_button = self.ui.create_button(
            operation_controls,
            text="連続購入モード開始",
            command=self.start_continuous_mode,
            state="disabled",
            fg_color="#4CAF50",  # 緑色
            hover_color="#45a049"
        )
        self.continuous_mode_button.grid(
            row=3, column=0, columnspan=2, padx=10, pady=5, sticky="ew")

        # 監視機能のフレームを追加
        monitor_frame = self.ui.create_frame(operation_controls)
        monitor_frame.grid(row=4, column=0, columnspan=2,
                           padx=10, pady=5, sticky="ew")

        # 監視開始ボタン
        self.start_monitor_button = self.ui.create_button(
            monitor_frame,
            text="商品更新監視開始",
            command=self.start_page_monitoring,
            state="disabled",
            fg_color="#2196F3"  # 青色
        )
        self.start_monitor_button.grid(
            row=0, column=0, padx=5, pady=5, sticky="ew")

        # 監視停止ボタン
        self.stop_monitor_button = self.ui.create_button(
            monitor_frame,
            text="監視停止",
            command=self.stop_page_monitoring,
            state="disabled",
            fg_color="#E53935",  # 赤色に変更
            hover_color="#C62828",  # 濃い赤（ホバー時）
            text_color="white"   # 文字色を白に
        )
        self.stop_monitor_button.grid(
            row=0, column=1, padx=5, pady=5, sticky="ew")

        # グリッド設定
        monitor_frame.grid_columnconfigure(0, weight=1)
        monitor_frame.grid_columnconfigure(1, weight=1)

    def create_status_section(self):
        """ステータス表示セクションの作成"""
        status_frame = self.ui.create_frame(self.main_frame)
        status_frame.pack(fill=tk.X, padx=10, pady=5)

        self.ui.create_label(
            status_frame,
            text="現在のステータス",
            font=self.header_font
        ).pack(anchor="w", padx=10, pady=5)

        # ステータスメッセージ表示
        self.status_message = self.ui.create_label(
            status_frame,
            text="待機中...",
            corner_radius=8,
            fg_color=("#E8F0FE", "#363636"),  # 薄い青（ライトモード）と濃いグレー（ダークモード）
            text_color=("#333333", "#FFFFFF"),  # 濃いグレー（ライトモード）と白（ダークモード）
            padx=10,
            pady=10
        )
        self.status_message.pack(fill=tk.X, padx=10, pady=5)

    def create_settings_section(self):
        """設定セクションの作成"""
        settings_frame = self.ui.create_frame(self.main_frame)
        settings_frame.pack(fill=tk.X, padx=10, pady=5)

        self.ui.create_label(
            settings_frame,
            text="設定",
            font=self.header_font
        ).pack(anchor="w", padx=10, pady=5)

        settings_controls = self.ui.create_frame(settings_frame)
        settings_controls.pack(fill=tk.X, padx=10, pady=5)

        # パスワード設定
        self.ui.create_label(
            settings_controls,
            text="パスワード:"
        ).grid(row=0, column=0, padx=10, pady=5, sticky="w")

        self.password_entry = self.ui.create_entry(
            settings_controls,
            width=250,
            show="*",
            placeholder_text="パスワードを入力してください"
        )
        self.password_entry.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        if self.config.get('password'):
            self.password_entry.insert(0, self.config.get('password'))

        self.save_button = self.ui.create_button(
            settings_controls,
            text="保存",
            command=self.save_config,
            width=80
        )
        self.save_button.grid(row=0, column=2, padx=10, pady=5)

        # プロファイル選択
        self.ui.create_label(
            settings_controls,
            text="Chromeプロファイル:"
        ).grid(row=1, column=0, padx=10, pady=5, sticky="w")

        # 利用可能なプロファイルを取得
        profiles = self.get_available_chrome_profiles()
        selected_profile = self.config.get('profile_directory', 'Default')

        # コンボボックスを作成
        self.profile_combobox = self.ui.create_combobox(
            settings_controls,
            values=profiles,
            state="readonly",
            width=250
        )
        self.profile_combobox.grid(
            row=1, column=1, padx=10, pady=5, sticky="ew")

        # 現在の設定値を選択
        if selected_profile in profiles:
            self.profile_combobox.set(selected_profile)
        else:
            self.profile_combobox.set('Default')

        # 詳細ログオプション
        self.verbose_checkbox = self.ui.create_checkbox(
            settings_controls,
            text="詳細ログを表示",
            variable=self.verbose_var
        )
        self.verbose_checkbox.grid(
            row=2, column=0, columnspan=3, padx=10, pady=5, sticky="w")

        # 列の重み設定
        settings_controls.grid_columnconfigure(1, weight=1)

    def create_log_section(self):
        """ログセクションの作成"""
        log_frame = self.ui.create_frame(self.main_frame)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.ui.create_label(
            log_frame,
            text="ログ",
            font=self.header_font
        ).pack(anchor="w", padx=10, pady=5)

        # ログテキストエリア
        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            wrap=tk.WORD,
            height=10,
            font=("メイリオ", 10)
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.log_text.config(state=tk.DISABLED)

    def log(self, message):
        """ログメッセージを記録"""
        if hasattr(self, 'logger') and self.logger:
            self.logger.info(message)
        else:
            # ロガーがまだ初期化されていない場合
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            log_msg = f"[{timestamp}] {message}"
            print(log_msg)

            # GUIのログエリアがすでに初期化されている場合
            if hasattr(self, 'log_text') and self.log_text:
                try:
                    self.log_text.config(state=tk.NORMAL)
                    self.log_text.insert(tk.END, log_msg + "\n")
                    self.log_text.see(tk.END)
                    self.log_text.config(state=tk.DISABLED)
                except Exception as e:
                    print(f"ログウィジェット更新エラー: {str(e)}")

    def update_status(self, message, level="info"):
        """ステータスメッセージを更新する"""
        if hasattr(self, 'status_message') and self.status_message:
            # レベル別の色設定
            colors = {
                "info": ("#333333", "#FFFFFF"),    # 黒/白 (ライト/ダークモード)
                "success": ("#006400", "#00CC00"),  # 濃い緑/明るい緑
                "warning": ("#CC6600", "#FFAA33"),  # 濃いオレンジ/明るいオレンジ
                "error": ("#CC0000", "#FF6666")    # 濃い赤/明るい赤
            }

            color = colors.get(level, colors["info"])
            self.status_message.configure(text=message, text_color=color)

            # 非同期更新のため即時反映
            self.root.update_idletasks()

        # ログにも記録
        self.log(f"ステータス: {message}")

    def ui_update_wrapper(self, callback, delay=0):
        """UIスレッドでコールバックを実行するためのラッパー関数"""
        self.root.after(delay, callback)

    def reset_ui_state_after_task(self):
        """タスク完了後のUI状態をリセット"""
        # 完全リセットを呼び出す
        self.reset_automation_state()

    def reset_automation_state(self):
        """自動化関連の状態を完全にリセット"""
        # タスクの状態をリセット
        if hasattr(self, 'task'):
            self.task.stop_requested = False
            self.task.running = False

        # 自動化インスタンスの状態をリセット
        if hasattr(self, 'automation') and self.automation:
            if hasattr(self.automation, 'stop_requested'):
                self.automation.stop_requested = False

        # UI要素を更新
        self.stop_button.configure(text="処理を停止", state="disabled")
        self.start_single_button.configure(state="normal")
        self.start_list_button.configure(state="normal")
        self.continuous_mode_button.configure(state="normal")

        # 監視関連のUIを初期状態に戻す（商品更新は1日1回なので、監視は再開しない）
        self.start_monitor_button.configure(state="normal")
        self.stop_monitor_button.configure(state="disabled")

        # タイマーをキャンセル
        if self.force_stop_timer:
            self.root.after_cancel(self.force_stop_timer)
            self.force_stop_timer = None

        # ステータスを更新
        self.update_status("状態をリセットしました。操作可能です。", "info")

    @gui_error_handler(operation="get_available_chrome_profiles")
    def get_available_chrome_profiles(self):
        """利用可能なChromeプロファイルの一覧を取得する"""
        profiles = ['Default']  # デフォルトプロファイルは常に含める

        # Chromeのユーザーデータフォルダパス
        user_data_dir = os.path.join(os.environ.get(
            'LOCALAPPDATA'), 'Google', 'Chrome', 'User Data')

        if os.path.exists(user_data_dir):
            # フォルダ内を検索してProfileで始まるフォルダを見つける
            for item in os.listdir(user_data_dir):
                if os.path.isdir(os.path.join(user_data_dir, item)) and item.startswith('Profile '):
                    profiles.append(item)

            self.log(f"{len(profiles)}個のChromeプロファイルを検出しました")
        else:
            self.log("Chromeのユーザーデータディレクトリが見つかりません")

        return profiles

    @gui_error_handler(operation="save_config")
    def save_config(self):
        """設定を保存する"""
        # UIから設定を取得
        raw_password = self.password_entry.get()
        selected_profile = self.profile_combobox.get()

        # 設定を更新
        self.config['password'] = raw_password  # 保存時に暗号化される
        self.config['verbose_log'] = self.verbose_var.get()
        self.config['profile_directory'] = selected_profile

        # 設定を保存
        if self.config_manager.save(self.config):
            self.update_status(
                f"設定を保存しました（プロファイル: {selected_profile}）", "success")
            self.dialog.show_info("設定を保存しました")
        else:
            self.update_status("設定の保存に失敗しました", "error")
            self.dialog.show_error("設定の保存に失敗しました")

    def create_tooltip(self, widget, text):
        """ウィジェットにツールチップを追加する"""
        def show_tooltip(event=None):
            x = y = 0
            if event:
                x = event.x_root + 25
                y = event.y_root + 25
            else:
                x = widget.winfo_rootx() + 25
                y = widget.winfo_rooty() + 25

            # ツールチップウィンドウの作成
            tooltip = tk.Toplevel(widget)
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{x}+{y}")

            label = tk.Label(tooltip, text=text, background="#ffffe0",
                             relief="solid", borderwidth=1, font=("メイリオ", 9))
            label.pack()

            # 一定時間後に消える
            widget.tooltip = tooltip
            widget._timer_id = widget.after(3000, lambda: hide_tooltip())

        def hide_tooltip(event=None):
            if hasattr(widget, "tooltip") and widget.tooltip is not None:
                try:
                    widget.tooltip.destroy()
                    widget.tooltip = None
                except:
                    pass  # 既に破棄されている場合は無視

            if hasattr(widget, "_timer_id"):
                try:
                    widget.after_cancel(widget._timer_id)
                except:
                    pass  # タイマーが既にキャンセルされている場合は無視

        # イベントをバインド
        widget.bind("<Enter>", show_tooltip)
        widget.bind("<Leave>", hide_tooltip)

    def add_tooltips(self):
        """UIコンポーネントにツールチップを追加"""
        tooltips = {
            # 既存のツールチップ
            self.chrome_button: "Chromeブラウザを起動します。自動化の前に必ず実行してください。",
            self.start_single_button: "商品詳細ページを開いた状態から購入処理を開始します。",
            self.start_list_button: "商品一覧ページ(検索結果画面)から商品を選んで購入できます。\n高速で操作してください！",
            self.open_site_button: "マップカメラのウェブサイトを開きます。",
            self.stop_button: "処理を即時停止します。商品を購入しそびれた場合は迅速に停止してください。",
            self.password_entry: "マップカメラサイトのログインパスワードを入力してください。",
            self.verbose_checkbox: "詳細なログ情報を表示します。問題が発生した場合に役立ちます。",
            self.status_message: "現在の処理状況を示します。エラーや警告は色で区別されます。",
            self.continuous_mode_button: "商品一覧ページから次々と商品を購入できるモードを開始します。\n商品更新タイミングでの連続購入に最適です。",
            # 新しく追加したツールチップ
            self.start_monitor_button: "11:30頃の商品更新を自動的に監視します。更新を検出したら通知します。",
            self.stop_monitor_button: "商品更新の監視を停止します。"  # 最後の要素にはカンマをつけない
        }

        for widget, text in tooltips.items():
            self.create_tooltip(widget, text)

    @gui_error_handler(operation="exit_app")
    def exit_app(self):
        """アプリケーションを終了する"""
        if self.dialog.show_confirm("終了してもよろしいですか？"):
            # 実行中のタスクがあれば強制停止
            # 監視中なら停止
            if hasattr(self, 'is_monitoring') and self.is_monitoring and hasattr(self, 'automation'):
                try:
                    self.automation.stop_monitoring()
                except:
                    pass
            if hasattr(self, 'task') and self.task.running:
                self.task.force_stop()

            # WebDriverの終了処理を確実に行う
            if hasattr(self, 'automation') and self.automation:
                try:
                    # 明示的にドライバーを終了
                    self.automation.driver.quit()
                    self.log("WebDriverを正常に終了しました")
                except Exception as e:
                    self.log(f"WebDriver終了中にエラーが発生: {str(e)}")

            self.root.destroy()

    @gui_error_handler(operation="start_chrome_checker")
    def start_chrome_checker(self):
        """Chromeの状態を定期的に確認する（改良版）"""
        def check_chrome():
            try:
                # Chromeプロセスが存在するか確認
                chrome_running = self.process_manager.check_chrome_running(
                    self.config.get('debug_port'))

                if not chrome_running:
                    # Chromeが終了している場合
                    if self.chrome_running:  # 以前は実行中だった場合
                        self.chrome_running = False
                        self.chrome_status.configure(
                            text="停止中", text_color="red")
                        self.start_single_button.configure(state="disabled")
                        self.start_list_button.configure(state="disabled")
                        self.continuous_mode_button.configure(state="disabled")
                        self.start_monitor_button.configure(state="disabled")

                        # 実行中のタスクがある場合は強制停止
                        if self.task.running:
                            self.log("Chromeが閉じられたため、実行中のタスクを停止します")
                            self.force_stop_automation()

                        self.log("Chromeが閉じられたか、実行状態を確認できませんでした。")
                        self.update_status(
                            "Chromeが閉じられました。再起動してください。", "warning")

                        # 再起動ボタンを有効化
                        self.chrome_button.configure(state="normal")

                # 次の確認をスケジュール
                self.root.after(1000, check_chrome)  # 1秒ごとにチェック（より頻繁に）

            except Exception as e:
                self.log(f"Chrome状態確認エラー: {str(e)}")
                # エラーが発生しても状態チェックを継続
                self.root.after(2000, check_chrome)

        # 最初のチェックを開始
        self.root.after(1000, check_chrome)

    @gui_error_handler(operation="start_chrome")
    def start_chrome(self):
        """Chromeを起動する（マップカメラサイトは開かない）"""
        if self.chrome_running:
            self.log("Chromeは既に実行中です")
            return

        # 古いプロセスをクリーンアップ
        self.process_manager.cleanup_chrome_drivers()

        self.log("Chromeを起動しています...")
        self.update_status("Chromeを起動しています...", "info")

        chrome_path = self.config.get('chrome_path')
        profile_dir = self.config.get('profile_directory', 'Default')
        debug_port = self.config.get('debug_port')

        # プロファイル情報をログに出力
        self.log(f"使用するプロファイル: {profile_dir}")

        # Chromeを指定のプロファイルとデバッグポートで起動（URLを指定せず新しいタブページを表示）
        command = f'"{chrome_path}" --remote-debugging-port={debug_port} --profile-directory="{profile_dir}"'
        self.log(f"実行コマンド: {command}")

        # サブプロセスとして実行
        subprocess.Popen(command, shell=True)
        self.log("Chromeを起動しました")
        self.update_status("Chromeが起動しました", "success")

        self.chrome_running = True
        self.chrome_status.configure(text="実行中", text_color="green")

        # 自動化関連のボタンを有効化
        self.start_single_button.configure(state="normal")
        self.start_list_button.configure(state="normal")
        self.continuous_mode_button.configure(state="normal")
        self.start_monitor_button.configure(state="normal")  # 監視ボタンも有効化

        # Chromeの状態を定期的に確認するタイマーを開始
        self.start_chrome_checker()

    @gui_error_handler(operation="force_stop_automation")
    def force_stop_automation(self):
        """処理を強制的に停止する - 改良版"""
        if not self.task.running:
            return

        self.log("処理を強制停止します...")
        self.update_status("処理を強制停止しています...", "warning")

        # WebDriver終了前にフラグを設定して、その後の操作を防止
        if hasattr(self, 'automation') and self.automation:
            if hasattr(self.automation, 'stop_requested'):
                self.automation.stop_requested = True

            # 新たに追加: 終了中フラグを設定
            if hasattr(self.automation, 'is_shutting_down'):
                self.automation.is_shutting_down = True

        # タスクを強制停止
        self.task.force_stop()

        # 強制停止タイマーをキャンセル
        if self.force_stop_timer:
            self.root.after_cancel(self.force_stop_timer)
            self.force_stop_timer = None

    @gui_error_handler(operation="stop_automation")
    def stop_automation(self):
        """実行中の自動化処理を停止する"""
        if not self.task.running:
            return

        self.log("処理の停止をリクエストしています...")

        # タスクに停止をリクエスト
        if self.task.stop():
            # 停止ボタンの表示を更新
            self.stop_button.configure(text="停止処理中...", state="disabled")
            self.update_status("停止リクエストを送信しました。処理が停止するまでお待ちください...", "warning")

            # 一定時間後に強制停止するタイマーを設定
            if self.force_stop_timer:
                self.root.after_cancel(self.force_stop_timer)

            # 設定ファイルから強制停止タイムアウト時間を取得（デフォルト2秒）
            force_stop_timeout = self.config.get('force_stop_timeout', 2000)
            self.force_stop_timer = self.root.after(
                force_stop_timeout, self.force_stop_automation)

    @gui_error_handler(operation="initialize_automation")
    def initialize_automation(self):
        """自動化クラスを初期化する"""
        try:
            self.log("自動化クラスの初期化を開始します")

            # セッション検証 - Chromeが再起動された場合のエラー対策
            if self.automation is not None:
                # 状態のリセットを追加
                if hasattr(self.automation, 'stop_requested'):
                    self.automation.stop_requested = False

                try:
                    # 現在のChromeセッションが有効かチェック
                    current_url = self.automation.driver.current_url

                    # さらにセッションの有効性を確認
                    try:
                        # JavaScriptを実行できるかテスト
                        document_state = self.automation.driver.execute_script(
                            "return document.readyState")

                        # URL情報の検証
                        if not current_url or current_url == "about:blank":
                            raise Exception("無効なURLです")

                        # URLがマップカメラのドメインを含むか確認
                        if "mapcamera.com" not in current_url:
                            self.log("現在マップカメラのページではありません。ページを切り替えます...")
                            # マップカメラのドメインに移動
                            self.automation.driver.get(
                                "https://www.mapcamera.com/")
                            self.log("マップカメラサイトに移動しました")
                            time.sleep(0.5)  # 少し待機

                            # ページが正しく読み込まれたか確認
                            loaded_state = self.automation.driver.execute_script(
                                "return document.readyState")
                            if loaded_state != "complete":
                                self.log(f"ページの読み込みが完了していません: {loaded_state}")
                                raise Exception("ページの読み込みが完了していません")
                    except Exception as e:
                        # セッションは存在するが不安定な状態
                        self.log(f"セッションが不安定です: {str(e)}")
                        raise Exception("セッションが不安定な状態です")

                    # URLが取得できて他のチェックもパスしたら有効なセッション
                    self.log("既存の自動化クラスを再利用します")
                    self.update_status("自動化クラスを再利用中...", "info")

                    # パスワードが変更されている場合は更新
                    current_password = self.password_entry.get()
                    if current_password:
                        self.automation.set_password(current_password)

                    return self.automation

                except Exception as e:
                    # セッションが無効な場合はエラーが発生する
                    self.log(f"セッションエラーが発生したため、自動化クラスを再初期化します: {str(e)}")
                    self.update_status(
                        "Chromeセッションが再起動されたため再初期化します...", "warning")

                    try:
                        # 既存インスタンスのクリーンアップ
                        self.automation.cleanup()
                    except Exception as cleanup_error:
                        self.log(f"クリーンアップエラー: {str(cleanup_error)}")

                    # 既存インスタンスを破棄
                    self.automation = None

            # 新しい自動化インスタンスを初期化
            password = self.password_entry.get()
            verbose = self.verbose_var.get()

            if not password:
                self.log("パスワードが設定されていません")
                self.update_status("パスワードが設定されていません", "error")
                self.dialog.show_error("パスワードを設定してください", "password_empty")
                return None

            self.update_status("自動化クラスを初期化しています...", "info")

            # トレースバック用にMapCameraAutomationクラスをインポート
            try:
                from mapcamera_automation import MapCameraAutomation

                # 自分自身への参照を渡して初期化
                self.automation = MapCameraAutomation(
                    password, get_config_path(), verbose_log=verbose, gui_handler=self)

                # タスクマネージャーに自動化インスタンスを設定
                self.task.automation = self.automation

                self.log("自動化クラスの初期化が完了しました")
                return self.automation

            except Exception as e:
                import traceback
                self.log(f"自動化クラスの初期化に失敗しました: {str(e)}")
                self.log(traceback.format_exc())
                self.update_status(f"自動化クラスの初期化に失敗しました", "error")
                self.dialog.show_error(f"自動化クラスの初期化に失敗しました: {str(e)}")
                return None

        except Exception as e:
            import traceback
            self.log(f"自動化初期化中にエラーが発生しました: {str(e)}")
            self.log(traceback.format_exc())
            return None

    @gui_error_handler(operation="run_automation_task")
    def run_automation_task(self, task_func, *args):
        """自動化タスクをバックグラウンドで実行する"""
        # 既に実行中の場合は開始しない
        if self.task.running:
            self.dialog.show_info("既に処理が実行中です。\n完了または停止するまでお待ちください。")
            return

        # 自動化クラスの初期化
        self.automation = self.initialize_automation()
        if not self.automation:
            self.update_status("自動化クラスの初期化に失敗しました", "error")
            return

        # タスク実行のコールバック
        def on_task_complete(result):
            if result:
                self.update_status("処理が完了しました", "success")
            else:
                self.update_status("処理は完了しましたが、エラーが発生しました", "warning")

            # UI要素を更新
            self.stop_button.configure(text="処理を停止", state="disabled")
            self.start_single_button.configure(state="normal")
            self.start_list_button.configure(state="normal")
            self.continuous_mode_button.configure(state="normal")

            # 強制停止タイマーをキャンセル
            if self.force_stop_timer:
                self.root.after_cancel(self.force_stop_timer)
                self.force_stop_timer = None

        # UI要素の状態を更新
        self.stop_button.configure(text="処理を停止", state="normal")
        self.start_single_button.configure(state="disabled")
        self.start_list_button.configure(state="disabled")
        self.continuous_mode_button.configure(state="disabled")

        # タスクを実行
        self.task.run(task_func, args, on_complete=on_task_complete)

    @gui_error_handler(operation="start_from_product_page")
    def start_from_product_page(self):
        """商品ページから購入を開始する"""
        self.update_status("商品ページから購入処理を開始します...", "info")

        def run_task():
            # 自動化の実行
            return self.automation.start_automation()

        self.run_automation_task(run_task)

    @gui_error_handler(operation="start_from_product_list")
    def start_from_product_list(self):
        """商品一覧から購入を開始する"""
        self.update_status("商品一覧から購入処理を開始します...", "info")

        def run_task():
            # タブの確認
            if not self.automation.find_best_tab():
                self.update_status("マップカメラのタブが見つかりません", "error")
                self.dialog.show_error("マップカメラのタブが見つかりません", "tab_not_found")
                return False

            # 商品クリック待機
            result = self.automation.wait_for_product_click()
            if result:
                self.update_status("商品が選択されました。購入処理を続行します...", "info")
                return self.automation.start_automation()
            else:
                return False

        self.run_automation_task(run_task)

    def start_continuous_mode(self):
        """連続購入モードを開始する"""
        # 状態のリセットと準備
        if hasattr(self, 'automation') and self.automation:
            if hasattr(self.automation, 'stop_requested'):
                self.automation.stop_requested = False
            if hasattr(self.automation, 'prevent_tab_switch'):
                self.automation.prevent_tab_switch = False

            # タブ参照をリセット
            for tab_attr in ['list_tab', 'product_tab']:
                if hasattr(self.automation, tab_attr):
                    delattr(self.automation, tab_attr)

        # 監視中かどうかの状態を保存
        was_monitoring = hasattr(self, 'is_monitoring') and self.is_monitoring

        # 監視中なら停止（商品更新は1日1回しかないため、再開する必要はない）
        if was_monitoring:
            if hasattr(self, 'automation') and self.automation:
                self.automation.stop_monitoring()
                self.is_monitoring = False
                self.update_status("連続購入モードのため、監視を停止しました", "info")

        # UI要素の状態を更新
        self.stop_button.configure(text="処理を停止", state="normal")
        self.start_single_button.configure(state="disabled")
        self.start_list_button.configure(state="disabled")
        self.continuous_mode_button.configure(state="disabled")

        # 監視ボタンを無効化
        self.start_monitor_button.configure(state="disabled")
        self.stop_monitor_button.configure(state="disabled")

        self.update_status("連続購入モードを開始します。商品一覧から購入したい商品を次々と選択できます。", "info")

        def run_task():
            # 自動化クラスを初期化
            self.automation = self.initialize_automation()
            if not self.automation:
                return False

            # 現在のタブが商品一覧ページかを確認
            current_tab = self.automation.driver.current_window_handle
            current_url = self.automation.driver.current_url

            # 最初に商品一覧タブを明確に記録（この参照を維持）
            if self.automation.is_product_list_page(current_url):
                self.automation.list_tab = current_tab
                self.log(f"商品一覧タブを記録しました: {current_tab}")
            else:
                # 商品一覧ページでなければ移動
                try:
                    self.automation.driver.get(
                        "https://www.mapcamera.com/search?sell=used&condition=other&sort=dateasc#result")
                    time.sleep(0.5)
                    if self.automation.is_product_list_page(self.automation.driver.current_url):
                        self.automation.list_tab = current_tab
                    else:
                        self.log("商品一覧ページへの移動に失敗しました")
                        return False
                except Exception as e:
                    self.log(f"ページ移動エラー: {str(e)}")
                    self.update_status("商品一覧ページへの移動に失敗しました。", "error")
                    return False

            # 連続購入ループ
            continuous_mode_active = True
            while continuous_mode_active and not self.task.stop_requested:
                try:
                    # ループの先頭でタブの有効性を確認
                    valid_tabs = self.automation.driver.window_handles

                    # list_tabの参照が無効になっていないかチェック
                    if hasattr(self.automation, 'list_tab') and self.automation.list_tab not in valid_tabs:
                        self.log("商品一覧タブが閉じられました。新しいタブを特定します。")
                        # 現在のタブが商品一覧ページかチェック
                        try:
                            current_tab = self.automation.driver.current_window_handle
                            current_url = self.automation.driver.current_url

                            if self.automation.is_product_list_page(current_url):
                                # 現在のタブが商品一覧ページなら、それをlist_tabとして記録
                                self.automation.list_tab = current_tab
                                self.log("商品一覧タブの参照を更新しました")
                            else:
                                # 商品一覧タブが見つからない場合は処理を継続
                                self.log("商品一覧タブが見つかりません。処理を継続します。")
                        except Exception as e:
                            self.log(f"タブ参照更新中にエラー: {str(e)}")

                    # 商品クリックを待機
                    if self.automation.wait_for_product_click():
                        # 商品が選択された場合

                        # SOLD OUT検出を先に行う
                        is_sold_out = self.automation.is_sold_out()

                        if is_sold_out:
                            # シンプルな通知のみを表示（ダイアログは表示しない）
                            self.update_status(
                                "この商品はSOLD OUTです。商品一覧タブに戻って次の商品を選択してください。", "warning")
                            self.log("SOLD OUT商品が検出されました。手動で商品一覧タブに切り替えてください。")

                            # タブ操作は一切行わない（切り替えも閉じるも行わない）

                            # このタブでの処理を完了し、次のループへ
                            continue
                        # SOLD OUTでない場合は購入処理を実行
                        result = self.automation.start_automation()

                        if result:
                            # 購入完了メッセージを表示
                            self.update_status(
                                "購入処理が完了しました。最終確認画面で停止しています。", "success")
                            self.log(
                                "最終確認画面で停止しています。注文確定後、次の商品を選択するには商品一覧タブに切り替えてください。")

                            # 遅延を追加して、ユーザーが最終確認画面を確認する時間を確保
                            time.sleep(2)  # 2秒間の待機

                            # 自動タブ切り替えが有効な場合のみ実行（現在はデフォルトで無効）
                            if self.config.get('auto_switch_tab', False):
                                try:
                                    if hasattr(self.automation, 'list_tab'):
                                        # 商品一覧タブに切り替え
                                        self.automation.driver.switch_to.window(
                                            self.automation.list_tab)
                                        self.automation.driver.execute_script(
                                            "window.scrollTo(0, 0);")
                                        self.update_status(
                                            "商品一覧タブに切り替えました。次の商品を選択してください。", "info")
                                    else:
                                        self.update_status(
                                            "商品一覧タブに手動で切り替えて、次の商品を選択してください。", "info")
                                except Exception as e:
                                    self.log(f"タブ切り替えエラー: {str(e)}")
                                    self.update_status(
                                        "タブ切り替えに失敗しました。手動で商品一覧タブに移動してください。", "warning")
                            else:
                                # 自動タブ切り替えが無効の場合のガイダンス
                                self.log(
                                    "最終確認画面で停止しています。注文確定後、次の商品を選択するには商品一覧タブに切り替えてください。")
                        else:
                            # 失敗した場合
                            self.update_status(
                                "購入処理が失敗しました。連続購入モードを終了します。", "warning")
                            continuous_mode_active = False
                    else:
                        # wait_for_product_clickがFalseを返した場合（停止要求など）
                        self.update_status("連続購入モードを終了します。", "info")
                        continuous_mode_active = False
                except Exception as e:
                    self.log(f"連続購入モードでエラーが発生: {str(e)}")

                    # エラーが発生した場合も確認ダイアログを表示
                    if self.dialog.show_confirm("エラーが発生しました。商品一覧ページに戻りますか？\n\n「いいえ」を選択すると現在のページにとどまります。"):
                        try:
                            # 商品一覧タブがある場合はそれに切り替え
                            if hasattr(self.automation, 'list_tab'):
                                self.automation.driver.switch_to.window(
                                    self.automation.list_tab)
                                self.update_status("商品一覧タブに切り替えました。", "info")
                            else:
                                # なければ通常のバック
                                if not self.automation.go_back_to_product_list():
                                    self.update_status(
                                        "商品一覧に戻れませんでした。連続購入モードを終了します。", "error")
                                    continuous_mode_active = False
                        except:
                            continuous_mode_active = False
                    else:
                        self.update_status(
                            "現在のページにとどまります。連続購入モードを終了します。", "info")
                        continuous_mode_active = False

                    # 短い待機を入れてエラー状態からの回復を試みる
                    time.sleep(1)

            # 終了処理
            try:
                if self.automation:
                    self.automation.cleanup()
            except Exception as e:
                self.log(f"クリーンアップエラー: {str(e)}")

            return True

        # タスクを開始
        self.task.run(run_task)

    @gui_error_handler(operation="start_page_monitoring")
    def start_page_monitoring(self):
        """商品更新の監視を開始 - 既存タブを使用するよう修正"""
        if not self.chrome_running:
            self.dialog.show_error("Chromeが起動していません。先にChromeを起動してください。")
            return

        # 既に監視中なら確実に停止してから開始する
        if hasattr(self, 'is_monitoring') and self.is_monitoring:
            # 既存の監視を完全に停止
            if hasattr(self, 'automation') and self.automation:
                self.automation.stop_monitoring()
                # 停止が完了するまで少し待機
                time.sleep(0.5)

        # 監視スレッド参照を明示的に確認
        if hasattr(self, 'automation') and hasattr(self.automation, 'monitor_thread'):
            try:
                # 既存のスレッドが生きていれば終了を待機
                if self.automation.monitor_thread and self.automation.monitor_thread.is_alive():
                    self.automation.monitor_thread.join(timeout=1.0)
            except Exception as e:
                print(f"監視スレッド終了待機エラー: {str(e)}")

        # 自動化クラスの初期化
        self.automation = self.initialize_automation()
        if not self.automation:
            self.update_status("自動化クラスの初期化に失敗しました", "error")
            return

        # 更新検出時のコールバック関数
        def on_update_detected():
            # 通知音を鳴らす（複数回鳴らして注意を引く）
            import winsound
            for _ in range(3):
                winsound.Beep(880, 300)  # 880Hzの音を300ミリ秒鳴らす
                time.sleep(0.2)

            # ステータス更新
            self.update_status("商品が更新されました！購入処理を開始できます。", "success")

            # ダイアログで通知
            self.dialog.show_info(
                "商品の更新を検出しました！\n「連続購入モード開始」ボタンをクリックして購入を開始できます。")

            # 連続購入ボタンを目立たせる
            self.continuous_mode_button.configure(
                fg_color="#FF4081",  # 目立つピンク色
                text="【更新検出】連続購入モード開始"
            )

        # 監視開始 - URLはNoneを渡し、既存タブを優先使用
        # フォールバックとしてデフォルトURLを指定
        default_url = "https://www.mapcamera.com/search?sell=used&condition=other&sort=dateasc#result"
        result = self.automation.monitor_page_updates(
            url=default_url, callback=on_update_detected)

        if result:
            self.is_monitoring = True
            # 設定から監視間隔を取得して表示
            interval = self.automation.config.get("monitoring_interval", 10)
            self.update_status(f"商品更新の監視を開始しました（{interval}秒間隔）", "info")

            # ボタン状態の更新
            self.start_monitor_button.configure(state="disabled")
            self.stop_monitor_button.configure(state="normal")
        else:
            self.update_status("監視の開始に失敗しました", "error")

    @gui_error_handler(operation="stop_page_monitoring")
    def stop_page_monitoring(self):
        """商品更新の監視を停止（非同期版）"""
        if not hasattr(self, 'is_monitoring') or not self.is_monitoring:
            return

        if not hasattr(self, 'automation') or not self.automation:
            self.is_monitoring = False
            return

        # 停止処理を開始
        self.update_status("監視を停止しています...", "info")

        # 非同期モードで監視停止をリクエスト
        result = self.automation.stop_monitoring(async_mode=True)

        if result:
            # 即時UIを更新
            self.is_monitoring = False
            self.update_status("監視停止を受け付けました", "info")

            # ボタン状態の更新
            self.start_monitor_button.configure(state="normal")
            self.stop_monitor_button.configure(state="disabled")

            # 連続購入ボタンを元に戻す
            self.continuous_mode_button.configure(
                fg_color="#4CAF50",  # 元の緑色
                text="連続購入モード開始"
            )

            # 1秒後に完了メッセージを表示（この時点ではスレッドはまだ終了処理中かもしれない）
            self.root.after(
                1000, lambda: self.update_status("監視を停止しました", "info"))
        else:
            self.update_status("監視の停止に失敗しました", "warning")

    @gui_error_handler(operation="open_mapcamera_site")
    def open_mapcamera_site(self):
        """マップカメラのウェブサイトを開く"""
        # URLを中古商品検索結果ページに変更
        mapcamera_url = "https://www.mapcamera.com/search?sell=used&condition=other&sort=dateasc#result"

        # すでにChromeが実行中の場合はそのままサイトを開く
        if self.chrome_running:
            webbrowser.open(mapcamera_url)
            self.log("マップカメラのウェブサイトを開きました")
            self.update_status("マップカメラサイトを開きました", "success")
            return

        # Chromeが実行中でない場合は先にChromeを起動する
        self.log("Chromeを起動し、マップカメラサイトを開きます...")
        self.update_status("Chromeを起動し、マップカメラサイトを開きます...", "info")

        # Chromeのパラメータ設定
        chrome_path = self.config.get('chrome_path')
        profile_dir = self.config.get('profile_directory', 'Default')
        debug_port = self.config.get('debug_port')

        # プロファイル情報をログに出力
        self.log(f"使用するプロファイル: {profile_dir}")

        # Chromeを指定のプロファイルとデバッグポートで起動
        command = f'"{chrome_path}" --remote-debugging-port={debug_port} --profile-directory="{profile_dir}" "{mapcamera_url}"'
        self.log(f"実行コマンド: {command}")

        # サブプロセスとして実行
        subprocess.Popen(command, shell=True)
        self.log("Chromeを起動し、マップカメラのウェブサイトを開きました")
        self.update_status("Chromeを起動し、マップカメラサイトを開きました", "success")

        # ステータスを更新
        self.chrome_running = True
        self.chrome_status.configure(text="実行中", text_color="green")

        # 自動化関連のボタンを有効化
        self.start_single_button.configure(state="normal")
        self.start_list_button.configure(state="normal")
        self.continuous_mode_button.configure(state="normal")

        # Chromeの状態を定期的に確認するタイマーを開始
        self.start_chrome_checker()


# メインアプリケーションの実行
if __name__ == '__main__':
    app = None
    try:
        # GUIの起動
        app = MapCameraGUI()
        app.root.mainloop()
    except Exception as e:
        print(f"予期せぬエラーが発生しました: {str(e)}")
        import traceback
        print(traceback.format_exc())
        # アプリケーションがクラッシュしても、WebDriverを終了させる
        if app and hasattr(app, 'automation') and app.automation:
            try:
                app.automation.driver.quit()
                print("WebDriverを終了しました")
            except:
                pass
    finally:
        # プログラム終了時の最終チェック
        if app and hasattr(app, 'automation') and app.automation:
            try:
                app.automation.driver.quit()
            except:
                pass
