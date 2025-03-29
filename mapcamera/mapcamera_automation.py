from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException, StaleElementReferenceException, InvalidSessionIdException
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
import sys
import re
import json
from datetime import datetime
import functools
import threading


# エラーメッセージの出力を抑制
os.environ['WDM_LOG_LEVEL'] = '0'
os.environ['WDM_PRINT_FIRST_LINE'] = 'False'

# 文字化け対策
if hasattr(sys, 'stdout') and sys.stdout is not None:
    if hasattr(sys.stdout, 'encoding') and sys.stdout.encoding != 'utf-8':
        if hasattr(sys.stdout, 'reconfigure'):
            sys.stdout.reconfigure(encoding='utf-8')


def error_handler(operation=None, include_url=True):
    """指定された操作のエラーを捕捉し、ログに記録するデコレータ（セッション無効処理強化版）"""

    def decorator(func):

        @functools.wraps(func)
        def wrapper(self, *args, **kwargs):
            try:
                # セッションが無効になっていないか確認
                if hasattr(self, 'driver'):
                    try:
                        # 簡単なコマンドでセッションをテスト
                        self.driver.title
                    except (InvalidSessionIdException, WebDriverException) as session_error:
                        # セッションが無効になっている場合
                        print("WebDriverセッションが無効になっているため操作をスキップします")
                        self.log_error(f"{operation or func.__name__}をスキップ: セッションが無効です",
                                       session_error, operation=operation, include_url=False)
                        # 終了中フラグを設定して後続の操作も中止
                        self.is_shutting_down = True
                        # ステータス更新を試みる
                        if hasattr(self, 'update_status'):
                            self.update_status(
                                "ブラウザセッションが終了しました。再起動してください。", "warning")
                        return False

                return func(self, *args, **kwargs)
            except Exception as e:
                op_name = operation or func.__name__
                self.log_error(f"{op_name}でエラーが発生",
                               e,
                               operation=op_name,
                               include_url=include_url)
                return False

        return wrapper

    return decorator


class MapCameraAutomation:

    def __init__(self, password, config_file=None, verbose_log=False, gui_handler=None):
        """マップカメラ自動化クラスの初期化"""
        try:
            print("MapCameraAutomationの初期化を開始します...")
            self.verbose_log = verbose_log  # 詳細ログフラグ
            self.gui_handler = gui_handler  # GUIハンドラへの参照
            self.is_shutting_down = False   # 追加: 終了中フラグを初期化
            self.initialize_driver()
            self.password = password
            self.config = self.load_config(config_file)
            print("待機時間を設定中...")

            # 高負荷環境向けに最適化されたタイムアウト設定とポーリング間隔
            self.wait = WebDriverWait(
                self.driver, 5, poll_frequency=0.2)  # 待機時間5秒、ポーリング間隔0.2秒
            self.driver.set_page_load_timeout(self.config.get(
                "page_load_timeout", 10))  # 設定ファイルの値を使用、デフォルトは10
            self.driver.set_script_timeout(self.config.get(
                "script_timeout", 10))  # 設定ファイルの値を使用、デフォルトは10

            # 初期化時に優先タブを探して設定
            self.find_best_tab()

            # 停止フラグの初期化
            self.stop_requested = False
            # タブ切り替え防止フラグを追加
            self.prevent_tab_switch = False

            print("初期化が完了しました")
            self.update_status("初期化が完了しました", "success")
        except Exception as e:
            error_msg = f"初期化エラー: {str(e)} ({type(e).__name__})"
            print(error_msg)
            if self.gui_handler:
                self.gui_handler.log(error_msg)
            # 初期化時はlog_errorメソッドがまだ使えないため、シンプルなエラー処理
            raise

    def log(self, message):
        """簡易ログ機能（update_statusとprintの組み合わせ）"""
        print(message)
        if hasattr(self, 'update_status'):
            self.update_status(message, "info")

    def load_config(self, config_file=None):
        """設定ファイルを読み込む"""
        default_config = {
            "wait_time": 5,
            "max_retries": 5,  # リトライ回数を3から5に増加
            "payment_method": "daibiki",  # 代金引換
            "debug_mode": False,
            "poll_frequency": 0.2,  # ポーリング間隔のデフォルト値
            "monitoring_interval": 10  # 監視間隔（秒）を追加
        }

        if config_file and os.path.exists(config_file):
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    user_config = json.load(f)
                    default_config.update(user_config)
                print(f"設定ファイルを読み込みました: {config_file}")
            except Exception as e:
                print(f"設定ファイルの読み込みに失敗しました: {str(e)}")

        return default_config

    def initialize_driver(self):
        """Chromeドライバーを初期化する"""
        try:
            print("ドライバーの初期化を開始します...")
            chrome_options = Options()

            # プロファイルの設定
            print("Chromeオプションを設定中...")
            # ハードコードされた値を使わず、システムの一般的なパスを使用
            chrome_options.add_argument(
                '--user-data-dir=C:\\Users\\' + os.getenv('USERNAME') +
                '\\AppData\\Local\\Google\\Chrome\\User Data')
            chrome_options.add_argument(
                '--profile-directory=Profile 7')  # Profile 7プロファイルを使用

            # デバッグポートを開く
            print("デバッグポート設定中...")
            chrome_options.add_experimental_option("debuggerAddress",
                                                   "127.0.0.1:9222")

            # ログレベルの設定
            chrome_options.add_argument('--log-level=3')
            chrome_options.add_argument('--silent')

            print("ChromeDriverのサービスを初期化中...")
            service = Service(ChromeDriverManager().install())

            # Windows環境でのみCREATE_NO_WINDOWフラグを設定
            if os.name == 'nt':  # Windowsの場合
                service.creation_flags = 0x08000000  # CREATE_NO_WINDOW

            print("WebDriverを作成中...")
            self.driver = webdriver.Chrome(service=service,
                                           options=chrome_options)

            print("ドライバーの初期化が完了しました")
        except Exception as e:
            error_msg = f"ドライバーの初期化エラー: {str(e)} ({type(e).__name__})"
            print(error_msg)
            # 初期化時はlog_errorメソッドが使えない可能性があるため、シンプルなエラー処理
            raise

    def update_status(self, message, level="info"):
        """GUI側にステータスを更新する"""
        # GUIハンドラが設定されている場合は、そちらに表示を委譲
        if self.gui_handler:
            if hasattr(self.gui_handler, "update_status"):
                self.gui_handler.update_status(message, level)
            elif hasattr(self.gui_handler, "log"):
                self.gui_handler.log(f"ステータス: {message}")

        # 常にコンソールにも表示
        print(message)

    def log_error(self, message, error, operation=None, include_url=True):
        """詳細なエラー情報をログに記録する"""
        error_type = type(error).__name__

        # 基本情報をログに記録
        error_message = f"{message}: {str(error)} ({error_type})"

        # 現在の操作情報を追加
        if operation:
            error_message += f" - 実行中の操作: {operation}"

        # 現在のURLを追加
        if include_url:
            try:
                current_url = self.driver.current_url
                error_message += f" - URL: {current_url}"
            except:
                error_message += " - URL: 取得不可"

        # コンソールに出力
        print(error_message)

        # GUIハンドラーがあればそちらにも表示
        if self.gui_handler:
            self.gui_handler.log(error_message)

        # 詳細ログが有効なら追加情報も表示
        if self.verbose_log:
            import traceback
            trace_info = traceback.format_exc()
            print(f"詳細なスタックトレース:\n{trace_info}")
            if self.gui_handler:
                self.gui_handler.log(f"詳細なスタックトレース:\n{trace_info}")

    def is_session_valid(self):
        """WebDriverセッションが有効かどうかを確認する"""
        try:
            if not hasattr(self, 'driver'):
                return False

            # 軽量なコマンドを実行してセッションをテスト
            self.driver.current_url
            return True
        except:
            return False

    def check_session_and_notify(self):
        """セッションの状態を確認し、無効な場合は通知する"""
        try:
            if not self.is_session_valid():
                print("WebDriverセッションが無効になっています")
                # 終了中フラグを設定
                self.is_shutting_down = True
                # 停止フラグも設定（他のメソッドがこれを見ている可能性がある）
                if hasattr(self, 'stop_requested'):
                    self.stop_requested = True

                if hasattr(self, 'update_status'):
                    self.update_status("ブラウザセッションが終了しました。再起動が必要です。", "warning")
                return False
            return True
        except Exception as e:
            print(f"セッション状態確認中にエラー: {str(e)}")
            return False  # エラーの場合は安全のためFalseを返す

    def show_browser_message(self, message, duration=None):
        """ブラウザ上にメッセージを表示する（GUIへの表示が優先）"""
        # 主にGUI側に表示を委託し、ブラウザ上での表示は最小限に
        self.update_status(message)

        # GUIハンドラがなく、または特別な理由がある場合のみブラウザに表示
        if not self.gui_handler or duration == 0:
            try:
                if self.verbose_log:
                    print(f"ブラウザメッセージを表示: {message}")
                escaped_message = message.replace("'",
                                                  "\\'").replace("\n", "\\n")

                # duration_js変数を定義
                duration_js = "null" if duration is None else str(duration *
                                                                  1000)

                js_code = f"""
                    var messageDiv = document.createElement('div');
                    messageDiv.id = 'automation-message';
                    messageDiv.style.position = 'fixed';
                    messageDiv.style.top = '50%';
                    messageDiv.style.left = '50%';
                    messageDiv.style.transform = 'translate(-50%, -50%)';
                    messageDiv.style.backgroundColor = 'rgba(0, 0, 0, 0.9)';
                    messageDiv.style.color = 'white';
                    messageDiv.style.padding = '20px';
                    messageDiv.style.borderRadius = '10px';
                    messageDiv.style.zIndex = '9999999';
                    messageDiv.style.fontSize = '18px';
                    messageDiv.style.fontWeight = 'bold';
                    messageDiv.style.maxWidth = '80%';
                    messageDiv.style.textAlign = 'center';
                    messageDiv.innerHTML = '{escaped_message}';
                    
                    // 既存のメッセージを削除
                    var oldMsg = document.getElementById('automation-message');
                    if (oldMsg) oldMsg.remove();
                    
                    document.body.appendChild(messageDiv);
                    
                    // 指定時間後に自動的に消去（非同期）
                    if ({duration_js} !== null) {{
                        setTimeout(function() {{
                            var msg = document.getElementById('automation-message');
                            if (msg) msg.remove();
                        }}, {duration_js});
                    }}
                """
                # メッセージを表示（非同期実行）
                self.driver.execute_script(js_code)

            except Exception as e:
                self.log_error("メッセージ表示エラー",
                               e,
                               operation="show_browser_message",
                               include_url=False)

    def check_stop(self):
        """停止リクエストがあるかどうかを確認"""
        return self.stop_requested

    def request_stop(self):
        """停止をリクエスト"""
        self.stop_requested = True
        self.update_status("停止をリクエストしました。処理を終了します...", "warning")

    def wait_for_element_with_stop_check(self, selector, timeout=5):
        """停止チェック付きの要素待機（動的待機時間対応版）"""
        start_time = time.time()
        poll_interval = 0.1  # 初期ポーリング間隔（短め）
        attempt = 0

        while time.time() - start_time < timeout:
            # 停止リクエストがあれば即時終了
            if self.check_stop():
                return None

            try:
                # 現在のポーリング間隔で要素を探す
                element = WebDriverWait(self.driver, poll_interval).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                if element:
                    if self.verbose_log:
                        print(
                            f"要素が{time.time() - start_time:.2f}秒で見つかりました: {selector}"
                        )
                    return element
            except TimeoutException:
                # 要素が見つからなければポーリング間隔を徐々に長くする
                attempt += 1
                # 試行回数に応じてポーリング間隔を調整（最大0.5秒まで）
                poll_interval = min(0.5, 0.1 + (attempt * 0.05))
                continue
            except Exception as e:
                # その他のエラーが発生した場合
                self.log_error("要素待機中のエラー",
                               e,
                               operation="wait_for_element_with_stop_check")
                time.sleep(0.05)  # エラー時は短く待機

        # タイムアウト
        if self.verbose_log:
            print(f"要素が{timeout}秒以内に見つかりませんでした: {selector}")
        return None

    def wait_for_any_element(self, selectors, timeout=5):
        """複数のセレクタから最初に見つかる要素を待機（停止チェック付き）"""
        if isinstance(selectors, str):
            selectors = [selectors]

        start_time = time.time()
        while time.time() - start_time < timeout:
            # 停止チェック
            if self.check_stop():
                return None, None

            # 各セレクタをチェック
            for selector in selectors:
                try:
                    element = WebDriverWait(self.driver, 0.5).until(
                        EC.element_to_be_clickable(
                            (By.CSS_SELECTOR, selector)))
                    if element:
                        return element, selector
                except:
                    continue

            # 短い間隔で再試行
            time.sleep(0.05)

        return None, None

    def handle_element_action(self,
                              selector,
                              action,
                              value=None,
                              timeout=5,
                              retries=5):
        """要素に対する様々なアクションを統一的に処理"""
        if isinstance(selector, str):
            selectors = [selector]
        else:
            selectors = selector

        for attempt in range(retries):
            if self.verbose_log:
                print(f"アクション '{action}' を実行中... (試行 {attempt + 1}/{retries})")

            # 停止チェック
            if self.check_stop():
                self.update_status("処理を停止しました", "warning")
                return False

            # 要素を待機
            element, found_selector = self.wait_for_any_element(
                selectors, timeout)

            if element:
                try:
                    if action == "click":
                        self.driver.execute_script("arguments[0].click();",
                                                   element)
                    elif action == "input":
                        element.clear()
                        element.send_keys(value)
                    elif action == "select":
                        # 将来的にselect要素対応が必要な場合
                        pass

                    if self.verbose_log:
                        print(f"アクション '{action}' が成功しました")
                    return True
                except Exception as e:
                    if self.verbose_log:
                        print(f"アクション実行エラー: {str(e)}")

                    # 要素が古くなっているなどのエラーでリトライ
                    if attempt < retries - 1:
                        time.sleep(0.5)
                    continue

            # 要素が見つからなかった場合
            if attempt < retries - 1:
                if self.verbose_log:
                    print(
                        f"要素が見つかりませんでした。リトライします... ({attempt + 2}/{retries})")
                time.sleep(0.5)
            else:
                self.update_status(f"要素が見つかりませんでした: {selectors}", "error")
                return False

        return False

    def get_tab_info(self, tab_handle):
        """タブの情報を取得する"""
        try:
            # 現在のタブを保存
            current_handle = self.driver.current_window_handle

            # 指定されたタブに切り替え
            self.driver.switch_to.window(tab_handle)

            url = self.driver.current_url
            info = {
                'handle': tab_handle,
                'url': url,
                'is_product': self.is_product_page(url),
                'is_list': self.is_product_list_page(url),
                'is_cart': '/cart' in url,
                'title': self.driver.title
            }

            # もとのタブに戻る
            self.driver.switch_to.window(current_handle)

            return info
        except Exception as e:
            if self.verbose_log:
                print(f"タブ情報取得エラー: {str(e)}")
            # エラーが発生した場合は無効なタブとして扱う
            return {
                'handle': tab_handle,
                'url': '',
                'is_product': False,
                'is_list': False,
                'is_cart': False,
                'title': '',
                'invalid': True
            }

    @error_handler(operation="find_best_tab")
    def find_best_tab(self):
        """利用可能なタブから最適なマップカメラのタブを見つける（エラーハンドリング強化版）"""
        print("最適なマップカメラタブを探しています...")

        # 最大試行回数を設定
        max_attempts = 3
        for attempt in range(max_attempts):
            try:
                # 現在のハンドルを保存
                try:
                    current_handle = self.driver.current_window_handle
                except:
                    current_handle = None

                # 全タブの情報を取得
                tab_handles = self.driver.window_handles

                if not tab_handles:
                    print("利用可能なタブがありません")
                    # 一時停止して再試行
                    time.sleep(1)
                    continue

                print(f"{len(tab_handles)}個のタブが見つかりました")

                # 各タブの情報を取得
                valid_tabs = []
                mapcamera_tabs = []

                for handle in tab_handles:
                    try:
                        info = self.get_tab_info(handle)
                        if not info.get('invalid', False):
                            valid_tabs.append(info)
                            if 'mapcamera.com' in info['url']:
                                mapcamera_tabs.append(info)
                                if self.verbose_log:
                                    print(f"マップカメラのタブを見つけました: {info['url']}")
                    except Exception as e:
                        if self.verbose_log:
                            print(f"タブ情報取得中にエラー: {str(e)}")

                if not mapcamera_tabs:
                    print("マップカメラのタブが見つかりませんでした")
                    # 最終試行でない場合は再試行
                    if attempt < max_attempts - 1:
                        time.sleep(1)
                        continue

                    # 有効なタブに戻る
                    if valid_tabs and current_handle != valid_tabs[0]['handle']:
                        self.driver.switch_to.window(valid_tabs[0]['handle'])
                    return False

                # 優先順位: 商品詳細 > 商品一覧 > その他マップカメラページ
                product_tabs = [
                    tab for tab in mapcamera_tabs if tab['is_product']
                ]
                list_tabs = [tab for tab in mapcamera_tabs if tab['is_list']]

                # 最適なタブを選択
                if product_tabs:
                    best_tab = product_tabs[0]
                    print(f"商品詳細ページのタブを選択しました: {best_tab['url']}")
                elif list_tabs:
                    best_tab = list_tabs[0]
                    print(f"商品一覧ページのタブを選択しました: {best_tab['url']}")
                else:
                    best_tab = mapcamera_tabs[0]
                    print(f"マップカメラのタブを選択しました: {best_tab['url']}")

                # 選択したタブに切り替え
                if current_handle != best_tab['handle']:
                    self.driver.switch_to.window(best_tab['handle'])
                    if self.verbose_log:
                        print(f"タブを切り替えました: {self.driver.current_url}")

                return True

            except Exception as e:
                print(f"試行 {attempt+1}/{max_attempts} でエラーが発生: {str(e)}")
                if attempt < max_attempts - 1:
                    print("再試行します...")
                    time.sleep(1)
                else:
                    print("最大試行回数に達しました")
                    return False

    def focus_on_correct_tab(self):
        """マップカメラのタブに焦点を当てる（改良版）"""
        try:
            # 現在のタブがマップカメラかどうかを確認
            try:
                current_url = self.driver.current_url
                is_mapcamera = 'mapcamera.com' in current_url
            except:
                is_mapcamera = False

            if not is_mapcamera:
                return self.find_best_tab()

            return True
        except Exception as e:
            self.log_error("タブ制御でエラーが発生", e, operation="focus_on_correct_tab")
            # エラーが発生した場合は改めて最適タブを探す
            return self.find_best_tab()

    def is_product_list_page(self, url):
        """URLが商品一覧ページかどうかを判定"""
        try:
            if self.verbose_log:
                print(f"URLを確認中: {url}")
            is_list = 'search' in url and ('mapcamera.com/search' in url or
                                           'www.mapcamera.com/search' in url)
            if self.verbose_log:
                print(f"商品一覧ページ判定結果: {is_list}")
            return is_list
        except Exception:
            return False

    def is_product_page(self, url):
        """URLが商品詳細ページかどうかを判定"""
        try:
            if self.verbose_log:
                print(f"URLを確認中: {url}")
            is_product = bool(
                re.match(r'https://www\.mapcamera\.com/item/\d+', url))
            if self.verbose_log:
                print(f"商品詳細ページ判定結果: {is_product}")
            return is_product
        except Exception:
            return False

    def is_product_page_by_content(self):
        """ページ内容から商品詳細ページかどうかを判定"""
        try:
            # 商品詳細ページに特有の要素を確認
            cart_button = self.driver.find_elements(By.CSS_SELECTOR,
                                                    "input[name='cartPut']")
            product_title = self.driver.find_elements(
                By.CSS_SELECTOR, "h1.item_name, h1.product-title")

            is_product = len(cart_button) > 0 or len(product_title) > 0
            if self.verbose_log:
                print(f"コンテンツによる商品詳細ページ判定結果: {is_product}")
            return is_product
        except Exception as e:
            if self.verbose_log:
                print(f"ページ内容確認でエラー: {str(e)}")
            return False

    def is_sold_out(self):
        """商品がSOLD OUTかどうかを確認する"""
        try:
            # 方法1: soldoutクラスを持つ要素を探す
            soldout_elements = self.driver.find_elements(
                By.CSS_SELECTOR, "p.soldout")
            if soldout_elements and len(soldout_elements) > 0:
                if self.verbose_log:
                    print("soldoutクラスを持つ要素が見つかりました")
                return True

            # 方法2: "SOLD OUT"テキストを含む要素を探す
            sold_out_text_elements = self.driver.find_elements(
                By.XPATH, "//*[contains(text(), 'SOLD OUT')]")
            if sold_out_text_elements and len(sold_out_text_elements) > 0:
                if self.verbose_log:
                    print("'SOLD OUT'テキストを含む要素が見つかりました")
                return True

            # 方法3: カートボタンが無効化されているか確認
            cart_buttons = self.driver.find_elements(
                By.CSS_SELECTOR,
                "input[name='cartPut'], button.cart-button, a.add-to-cart")
            if not cart_buttons or len(cart_buttons) == 0:
                # カートボタンが見つからない場合、タイトルに「売約済」などが含まれているか確認
                title_element = self.driver.find_element(
                    By.CSS_SELECTOR, "h1.item_name, h1.product-title")
                if title_element and ("売約済" in title_element.text
                                      or "完売" in title_element.text):
                    if self.verbose_log:
                        print("商品タイトルに「売約済」または「完売」が含まれています")
                    return True

                if self.verbose_log:
                    print("カートボタンが見つかりません。商品ページの構造が変更されているか、SOLDOUTの可能性があります")
                # ここでは、カートボタンがないだけではSOLD OUTとは判断しない（構造変更の可能性もあるため）

            return False

        except Exception as e:
            if self.verbose_log:
                print(f"SOLD OUTチェック中にエラーが発生: {str(e)}")
            # エラーの場合は安全のためFalseを返す（機能が不完全でも既存の処理は続行できるように）
            return False

    def handle_sold_out(self):
        """SOLD OUT商品の処理"""
        try:
            if not self.focus_on_correct_tab():
                return False

            message = "この商品はSOLD OUTです。"
            print(message)
            self.update_status(message, "warning")

            # 商品一覧に自動で戻らず、True/Falseを返すだけに変更
            return True

        except Exception as e:
            self.log_error("SOLD OUT処理中にエラーが発生",
                           e,
                           operation="handle_sold_out")
            return False

    @error_handler(operation="go_back_to_product_list")
    def go_back_to_product_list(self):
        """最適な商品一覧ページに戻る"""
        # 最後に閲覧した商品一覧ページURLを記録する変数を追加（クラスのインスタンス変数として）
        if not hasattr(self, 'last_product_list_url'):
            self.last_product_list_url = "https://www.mapcamera.com/search?sell=used&condition=other&sort=dateasc#result"

        print(f"商品一覧ページに戻ります: {self.last_product_list_url}")
        self.update_status("商品一覧ページに移動します...", "info")

        # まず「戻る」ボタンで戻ってみる
        current_url = self.driver.current_url
        self.driver.back()
        time.sleep(0.5)

        # 戻った先が商品一覧ページかチェック
        if self.is_product_list_page(self.driver.current_url):
            print("「戻る」ボタンで商品一覧ページに戻りました")
            # 現在のURLを記録（次回のために）
            self.last_product_list_url = self.driver.current_url
            self.update_status("商品一覧ページに戻りました。次の商品を選択してください。", "info")

            # ページトップにスクロール
            self.driver.execute_script("window.scrollTo(0, 0);")

            return True

        # 戻れなかった場合は記録済みの商品一覧URLに直接移動
        self.driver.get(self.last_product_list_url)
        time.sleep(0.5)

        # 移動先が商品一覧ページかチェック
        if self.is_product_list_page(self.driver.current_url):
            print(f"直接URLで商品一覧ページに移動しました: {self.last_product_list_url}")
            self.update_status("商品一覧ページに移動しました。次の商品を選択してください。", "info")

            # ページトップにスクロール
            self.driver.execute_script("window.scrollTo(0, 0);")

            return True

        # それでもダメな場合はデフォルトの検索ページに移動
        self.driver.get("https://www.mapcamera.com/search")
        time.sleep(0.5)
        print("デフォルトの検索ページに移動しました")
        self.update_status("検索ページに移動しました。検索条件を設定してください。", "info")

        # ページトップにスクロール
        self.driver.execute_script("window.scrollTo(0, 0);")

        return True

    @error_handler(operation="continue_shopping")
    def continue_shopping(self):
        """購入処理完了後、新しいタブで商品一覧に移動して次の商品購入に備えるが、タブは切り替えない"""
        # 購入完了後の状態をリセット
        self.stop_requested = False

        # 最後に使用した商品一覧URL（なければデフォルト）
        search_url = getattr(
            self, 'last_product_list_url',
            "https://www.mapcamera.com/search?sell=used&condition=other&sort=dateasc#result"
        )

        print(f"バックグラウンドで商品一覧ページを開きます: {search_url}")
        self.update_status("バックグラウンドで商品一覧ページを準備しています...", "info")

        # 現在のタブを記憶
        current_tab = self.driver.current_window_handle

        # 新しいタブで商品一覧を開く
        self.driver.execute_script(f"window.open('{search_url}', '_blank');")

        # 新しく開いたタブのハンドルを取得
        tabs = self.driver.window_handles
        new_tab = tabs[-1]  # 最後に開いたタブ

        # 一時的に新しいタブに切り替えて情報を取得
        self.driver.switch_to.window(new_tab)

        # 新しいタブが商品一覧ページかチェック
        is_list_page = self.is_product_list_page(self.driver.current_url)

        # URLを記録（次回のために）
        if is_list_page:
            self.last_product_list_url = self.driver.current_url
            print(f"商品一覧ページのURLを記録しました: {self.driver.current_url}")

        # 元のタブに戻る
        self.driver.switch_to.window(current_tab)

        if is_list_page:
            self.update_status("バックグラウンドで商品一覧ページを準備しました。タブを切り替えて次の商品を選択できます。",
                               "info")
            return True
        else:
            self.update_status("バックグラウンドタブの準備に問題があります。", "warning")
            return False

    @error_handler(operation="wait_for_product_click")
    def wait_for_product_click(self):
        """商品一覧ページで商品クリックを待機し、新しいタブで開く - 改良版"""
        # 追加: 終了中チェック
        if hasattr(self, 'is_shutting_down') and self.is_shutting_down:
            print("シャットダウン中のため操作をスキップします")
            return False

        # 追加: セッション状態チェック
        if not self.check_session_and_notify():
            return False

        # 現在のページが商品一覧ページの場合のみタブ切り替え処理を実行
        if self.is_product_list_page(self.driver.current_url):
            # タブの有効性チェック
            valid_tabs = self.driver.window_handles

            # list_tabが設定されていて、まだ存在する場合のみ使用
            if hasattr(self, 'list_tab') and self.list_tab in valid_tabs:
                # list_tabに切り替え
                self.driver.switch_to.window(self.list_tab)
                # 商品一覧ページでのみステータスメッセージを更新
                self.update_status("商品一覧ページです。購入したい商品をクリックしてください。", "info")
            else:
                # 有効なタブがない場合は正しいタブを検索
                if not self.focus_on_correct_tab():
                    return False
                # 現在のタブを商品一覧タブとして記録
                current_tab = self.driver.current_window_handle
                if self.is_product_list_page(self.driver.current_url):
                    self.list_tab = current_tab
                    # 商品一覧ページでのみステータスメッセージを更新
                    self.update_status("商品一覧ページです。購入したい商品をクリックしてください。", "info")
        else:
            # 商品詳細ページや最終確認画面など、商品一覧ページでない場合は
            # タブ切り替えを行わず、単に現在のタブが有効かチェックするのみ
            if not self.focus_on_correct_tab():
                return False
            # ステータスメッセージは更新しない（前のメッセージを維持）

        print("商品クリックを待機中...")
        self.update_status("商品一覧ページです。購入したい商品をクリックしてください。", "info")

        # 現在のタブ情報を保存
        initial_tab = self.driver.current_window_handle
        initial_url = self.driver.current_url
        initial_tabs = self.driver.window_handles.copy()  # タブのリストをコピー

        # 商品一覧ページURLを記録
        if self.is_product_list_page(initial_url):
            self.last_product_list_url = initial_url
            if self.verbose_log:
                print(f"商品一覧ページのURLを記録しました: {initial_url}")

            # リストタブを記録
            self.list_tab = initial_tab

        # パッチング関数: 直接window.openを使用して新しいタブで商品を開く（修正版）
        self.driver.execute_script("""
            // 既存のコンテキストメニューイベントが正常に動作するようにする
            window._originalOpenHook = window.open;
            
            // マップカメラのドメインかどうかをチェックする関数
            function isMapCameraDomain() {
                return window.location.hostname.includes('mapcamera.com');
            }
            
            // 商品リンクのクリックを監視する関数
            function setupProductLinkWatcher() {
                // マップカメラのドメインでない場合は処理しない
                if (!isMapCameraDomain()) {
                    console.log("マップカメラ以外のドメインのため、リンク監視を行いません");
                    return 0;
                }
                
                var links = document.querySelectorAll('a[href*="/item/"]');
                var modifiedCount = 0;
                
                for (var i = 0; i < links.length; i++) {
                    var link = links[i];
                    var href = link.getAttribute('href');
                    
                    // 商品リンクのみを対象（/item/maker や /item/category を除外）
                    if (href.indexOf('/item/maker') >= 0 || 
                        href.indexOf('/item/category') >= 0 ||
                        href.indexOf('/item/list') >= 0) {
                        continue;
                    }
                    
                    // すでに処理済みのリンクはスキップ
                    if (link.getAttribute('data-modified') === 'true') {
                        continue;
                    }
                    
                    // オリジナルのクリックハンドラを保存
                    var originalClickHandler = link.onclick;
                    
                    // クリックイベントを修正
                    link.onclick = function(e) {
                        e.preventDefault();
                        e.stopPropagation();
                        
                        // 元のURLを保存
                        var targetUrl = this.href;
                        
                        // 新しいタブで開く
                        var newTab = window._originalOpenHook(targetUrl, '_blank');
                        
                        // イベントキャンセル
                        return false;
                    };
                    
                    // マーキング
                    link.setAttribute('data-modified', 'true');
                    link.setAttribute('target', '_blank');  // 念のため
                    modifiedCount++;
                }
                
                return modifiedCount;
            }
            
            // 定期的なドメインチェック機能を追加
            function checkDomainAndSetupWatcher() {
                try {
                    if (isMapCameraDomain()) {
                        return setupProductLinkWatcher();
                    } else {
                        console.log("マップカメラ以外のドメインです");
                        return -1; // マップカメラ以外を示す値を返す
                    }
                } catch (e) {
                    console.error("ドメインチェック中にエラー:", e);
                    return -2; // エラーを示す値を返す
                }
            }
            
            // 初回実行
            var result = checkDomainAndSetupWatcher();
            console.log("Modified " + result + " product links");
            
            // インターバルIDを保存（既存のものがあれば再利用）
            if (!window._productLinkWatcherInterval) {
                window._productLinkWatcherInterval = setInterval(function() {
                    var newCount = checkDomainAndSetupWatcher();
                    if (newCount > 0) {
                        console.log("Found and modified " + newCount + " new product links");
                    }
                }, 2000);
            }
            
            return result;
        """)

        print("商品リンクを処理しました。クリックされるのを待機中...")

        # 停止チェックの間隔を短く設定
        stop_check_interval = 0.05  # 50ミリ秒ごとに停止チェック
        last_stop_check = time.time()
        last_message_update = time.time()
        last_domain_check = time.time()
        domain_check_interval = 1.0  # 1秒ごとにドメインをチェック

        # 新しいタブが開かれるまで待機
        while True:
            # 停止チェック
            current_time = time.time()
            if (current_time - last_stop_check) >= stop_check_interval:
                if self.check_stop():
                    print("ユーザーリクエストにより処理を停止します")
                    self.update_status("処理を停止しました", "warning")
                    return False
                last_stop_check = current_time

            # 終了中フラグのチェックを追加
            if hasattr(self, 'is_shutting_down') and self.is_shutting_down:
                print("シャットダウン中のため操作を停止します")
                return False

            # 定期的なドメインチェックを追加
            if (current_time - last_domain_check) >= domain_check_interval:
                try:
                    # 現在のタブが有効かチェック
                    if self.tab_exists(initial_tab):
                        # 現在アクティブなタブを保存
                        current_active_tab = self.driver.current_window_handle

                        # タブ切り替えを最小限に抑えるため、現在のタブが初期タブの場合のみ詳細チェック
                        if current_active_tab == initial_tab:
                            # セッションが有効かチェック
                            if not self.is_session_valid():
                                print("セッションが無効になりました。処理を中止します")
                                self.update_status(
                                    "ブラウザセッションが終了しました。再起動してください。", "warning")
                                return False

                            # 初期タブのドメインとURLを取得
                            current_url = self.driver.current_url
                            is_mapcamera = 'mapcamera.com' in current_url

                            if not is_mapcamera:
                                print("初期タブがマップカメラ以外のドメインに移動しました。商品クリック待機を終了します")
                                self.update_status(
                                    "マップカメラ以外のページに移動しました。購入処理を終了します。", "warning")
                                return False

                            # 商品一覧ページに戻ったかチェック
                            current_is_list_page = self.is_product_list_page(
                                current_url)

                            # 前回のURLと今回のURLを使って状態変化を検出
                            if hasattr(self, '_last_checked_url'):
                                previous_is_list_page = self.is_product_list_page(
                                    self._last_checked_url)
                                previous_url = self._last_checked_url

                                # 非一覧ページから一覧ページに変わった場合（復帰）
                                if current_is_list_page and not previous_is_list_page:
                                    print("商品一覧ページに戻りました。リンク変換スクリプトを再実行します")
                                    self.update_status(
                                        "商品一覧ページに戻りました。商品をクリックできます。", "info")
                                    self._apply_link_conversion_script()
                                    # リロードタイムスタンプを更新（復帰時は必ずスクリプト適用）
                                    self._last_reload_timestamp = int(
                                        time.time() * 1000)

                                # 商品一覧ページから別のページに移動した場合
                                elif not current_is_list_page and previous_is_list_page:
                                    # 購入処理中フラグがない場合のみメッセージを更新
                                    if not hasattr(self, 'purchase_in_progress') or not self.purchase_in_progress:
                                        print("商品一覧ページから別のページに移動しました")
                                        self.update_status(
                                            "商品一覧ページから別のページに移動しました。", "info")

                                # 同じ商品一覧ページ内での変化（リロードなど）
                                elif current_url == previous_url and current_is_list_page:
                                    # リロード頻度制限のためのタイムスタンプチェック
                                    current_time_millis = int(
                                        time.time() * 1000)
                                    if hasattr(self, '_last_reload_timestamp'):
                                        time_since_last_reload = current_time_millis - self._last_reload_timestamp
                                        # 5秒以内の再検出は無視（誤検出防止）
                                        if time_since_last_reload >= 5000:
                                            # 修正されたリンクをカウント（スクリプトの生存確認）
                                            link_count = self.driver.execute_script("""
                                                // クリックハンドラーが適用されたリンクの数をカウント
                                                var links = document.querySelectorAll('a[data-modified="true"]');
                                                return links.length;
                                            """)

                                            # リンクが0であれば、スクリプトが消えている可能性が高い
                                            if link_count == 0:
                                                print(
                                                    f"リンク変換が無効になっています。スクリプトを再適用します（検出リンク数: {link_count}）")
                                                self.update_status(
                                                    "リンク変換を再適用します", "info")
                                                self._apply_link_conversion_script()
                                                # リロードタイムスタンプを更新
                                                self._last_reload_timestamp = current_time_millis
                                    else:
                                        # 初回のリロードタイムスタンプを設定
                                        self._last_reload_timestamp = current_time_millis

                            # 現在のURLを保存して次回のチェックに使用
                            self._last_checked_url = current_url

                            # 商品一覧ページで適切なメッセージを表示（ただし前回と同じメッセージは表示しない）
                            if current_is_list_page:
                                # 前回のメッセージと比較して、同じなら表示しない
                                new_message = "商品一覧ページです。購入したい商品をクリックしてください。"
                                if not hasattr(self, '_last_status_message') or self._last_status_message != new_message:
                                    self.update_status(new_message, "info")
                                    self._last_status_message = new_message
                        else:
                            # 現在のタブが初期タブと異なる場合、初期タブの状態を調べるためでも切り替えない
                            # 完全に非侵入的な動作のためにタブ切り替えを行わない

                            # 購入処理中かどうかを確認
                            if hasattr(self, 'product_tab') and current_active_tab == self.product_tab:
                                # 商品タブがアクティブなら、ガイダンスを表示
                                if not hasattr(self, '_last_status_message') or self._last_status_message != "商品一覧ページに切り替えて、次の商品を選択できます。":
                                    self.update_status(
                                        "商品一覧ページに切り替えて、次の商品を選択できます。", "info")
                                    self._last_status_message = "商品一覧ページに切り替えて、次の商品を選択できます。"
                            else:
                                # その他のタブでは何もしない（ユーザーの自由な操作を尊重）
                                pass

                        last_domain_check = current_time
                    else:
                        print("初期タブが存在しません")
                        last_domain_check = current_time
                except Exception as e:
                    # エラーが発生した場合はセッションが無効になっている可能性が高い
                    print(f"ドメインチェック中にエラー: {str(e)}")
                    if "invalid session id" in str(e).lower():
                        print("セッションが無効になりました。処理を中止します")
                        self.update_status(
                            "ブラウザセッションが終了しました。再起動してください。", "warning")
                        return False
                    last_domain_check = current_time

            # 定期的にメッセージを更新
            if (current_time - last_message_update) > 3:
                # 安全にタブ存在チェック
                if self.tab_exists(initial_tab):
                    # 現在のタブがリストタブのままか確認
                    if self.driver.current_window_handle == initial_tab:
                        current_url = self.driver.current_url
                        if self.is_product_list_page(current_url):
                            self.update_status("商品一覧ページです。購入したい商品をクリックしてください。",
                                               "info")
                last_message_update = current_time

            # 新しいタブが開かれたかチェック
            try:
                current_tabs = self.driver.window_handles

                # タブが閉じられた場合でも新しいタブを検出できるようにする
                new_tabs = [
                    tab for tab in current_tabs if tab not in initial_tabs]
                if new_tabs:
                    new_tab = new_tabs[-1]  # 最新のタブを取得

                    # 念のため元のタブが有効か確認
                    if self.tab_exists(initial_tab):
                        if self.driver.current_window_handle != initial_tab:
                            # 安全にタブ切り替え
                            self.safe_switch_to_tab(initial_tab)

                    # 新しいタブに切り替え - 安全に
                    if not self.safe_switch_to_tab(new_tab):
                        # タブが閉じられていたら続行
                        continue

                    # 明示的なWaitを使用してページ読み込みを待機
                    try:
                        # URLがabout:blankから変わるまで待機（最大10秒）
                        WebDriverWait(self.driver, 10, poll_frequency=0.2).until(
                            lambda d: d.current_url != "about:blank"
                        )

                        # さらにページ本体が読み込まれるのを待機（最大8秒）
                        WebDriverWait(self.driver, 8, poll_frequency=0.2).until(
                            lambda d: d.execute_script(
                                "return document.readyState") != "loading"
                        )

                        if self.verbose_log:
                            print(f"ページが読み込まれました: {self.driver.current_url}")
                    except Exception as e:
                        print(f"ページ読み込み待機中にエラー: {str(e)}")
                        # エラーが発生しても処理を継続

                    # マップカメラのドメインであることを確認（修正版）
                    current_url = self.driver.current_url
                    if 'mapcamera.com' not in current_url:
                        print(f"新しいタブがマップカメラのドメインではありません: {current_url}")

                        # 現在のタブ数をチェック
                        current_tab_count = len(self.driver.window_handles)
                        if current_tab_count <= 1:
                            # タブが1つしかない場合は閉じずに商品一覧ページに戻す
                            print("タブが1つしかないため閉じずに商品一覧ページに戻します")
                            self.update_status(
                                "マップカメラ以外のページが検出されました。商品一覧ページに戻します。", "info")
                            try:
                                # マップカメラの商品一覧ページに戻る
                                self.driver.get(
                                    self.last_product_list_url or "https://www.mapcamera.com/search?sell=used&condition=other&sort=dateasc#result")
                                # ページ読み込み完了を待機
                                WebDriverWait(self.driver, 10, poll_frequency=0.2).until(
                                    lambda d: d.execute_script(
                                        "return document.readyState") == "complete"
                                )

                                # リンク変換スクリプトを適用
                                self._apply_link_conversion_script()

                                # 初期タブの情報を更新
                                initial_tab = self.driver.current_window_handle
                                initial_url = self.driver.current_url
                                initial_tabs = self.driver.window_handles.copy()
                            except Exception as e:
                                print(f"商品一覧ページへの復帰中にエラー: {str(e)}")
                                if "invalid session id" in str(e).lower():
                                    print("セッションが無効になりました。処理を中止します")
                                    self.update_status(
                                        "ブラウザセッションが終了しました。再起動してください。", "warning")
                                    return False
                        else:
                            # 複数タブがある場合は自動的に閉じるのではなく、情報のみ表示
                            print("マップカメラ以外のページが開かれましたが、閉じずに監視を継続します")
                            self.update_status(
                                "マップカメラ以外のページが開かれています。監視は継続します。", "info")

                            # タブ切り替えをせず、現在のタブのままにする
                            # 元のタブに戻る処理を削除

                        # リストの更新（次の検出のため）
                        initial_tabs = current_tabs.copy()
                        continue

                    if self.is_product_page(
                            current_url) or self.is_product_page_by_content():
                        print(f"新しいタブで商品ページを検出: {current_url}")
                        self.update_status(
                            "商品ページを検出しました。購入処理を開始します...", "success")

                        # タブ情報を記録
                        self.product_tab = new_tab
                        self.list_tab = initial_tab

                        return True
                    else:
                        # 商品ページでない場合は閉じて元のタブに戻る
                        print(f"新しいタブが商品ページではありません: {current_url}")
                        self.update_status("商品ページではありません。", "warning")
                        self.driver.close()
                        # 安全にタブ切り替え
                        if self.tab_exists(initial_tab):
                            self.driver.switch_to.window(initial_tab)
                        # リストの更新（次の検出のため）
                        initial_tabs = current_tabs.copy()
            except Exception as e:
                print(f"タブチェック中にエラー: {str(e)}")
                # セッションが無効になっている場合は処理を中止
                if "invalid session id" in str(e).lower():
                    print("セッションが無効になりました。処理を中止します")
                    self.update_status(
                        "ブラウザセッションが終了しました。再起動してください。", "warning")
                    return False
                # 一時的なエラーの場合は少し待機して継続
                time.sleep(0.1)

            # 現在のタブが変わっていないか確認（商品が同じタブで開いてしまった場合）
            try:
                if self.tab_exists(
                        initial_tab
                ) and self.driver.current_window_handle == initial_tab:
                    current_url = self.driver.current_url
                    if current_url != initial_url and (
                            self.is_product_page(current_url)
                            or self.is_product_page_by_content()):
                        print("警告: 商品が同じタブで開かれました。別タブで再オープンします。")

                        # 新しいタブで商品一覧ページを開き、元のタブを商品ページとして使用
                        self.driver.execute_script(
                            f"window.open('{initial_url}', '_blank');")

                        # 明示的なWaitを使用してタブのロードを待機
                        time.sleep(0.1)  # 最小限の待機

                        # 新しく開いたタブをリストタブとして設定
                        new_tabs = [
                            tab for tab in self.driver.window_handles
                            if tab not in initial_tabs
                        ]
                        if new_tabs:
                            self.list_tab = new_tabs[-1]
                            self.product_tab = initial_tab

                            print(f"商品一覧タブを作成しました: {self.list_tab}")
                            self.update_status("商品ページを検出しました。購入処理を開始します...",
                                               "success")
                            return True
            except Exception as e:
                print(f"タブ状態チェック中にエラー: {str(e)}")
                # セッションが無効になっている場合は処理を中止
                if "invalid session id" in str(e).lower():
                    print("セッションが無効になりました。処理を中止します")
                    self.update_status(
                        "ブラウザセッションが終了しました。再起動してください。", "warning")
                    return False

            # 少し待機
            time.sleep(0.05)

    def _apply_link_conversion_script(self):
        """リンク変換スクリプトを適用するヘルパーメソッド"""
        try:
            # リンク変換スクリプトを完全に再実行
            result = self.driver.execute_script("""
                // 既存のコンテキストメニューイベントが正常に動作するようにする
                window._originalOpenHook = window.open;
                
                // 商品リンクのクリックを監視する関数
                function setupProductLinkWatcher() {
                    var links = document.querySelectorAll('a[href*="/item/"]');
                    var modifiedCount = 0;
                    
                    for (var i = 0; i < links.length; i++) {
                        var link = links[i];
                        var href = link.getAttribute('href');
                        
                        // 商品リンクのみを対象（/item/maker や /item/category を除外）
                        if (href.indexOf('/item/maker') >= 0 || 
                            href.indexOf('/item/category') >= 0 ||
                            href.indexOf('/item/list') >= 0) {
                            continue;
                        }
                        
                        // すでに処理済みのリンクはスキップ
                        if (link.getAttribute('data-modified') === 'true') {
                            continue;
                        }
                        
                        // オリジナルのクリックハンドラを保存
                        var originalClickHandler = link.onclick;
                        
                        // クリックイベントを修正
                        link.onclick = function(e) {
                            e.preventDefault();
                            e.stopPropagation();
                            
                            // 元のURLを保存
                            var targetUrl = this.href;
                            
                            // 新しいタブで開く
                            var newTab = window._originalOpenHook(targetUrl, '_blank');
                            
                            // イベントキャンセル
                            return false;
                        };
                        
                        // マーキング
                        link.setAttribute('data-modified', 'true');
                        link.setAttribute('target', '_blank');  // 念のため
                        modifiedCount++;
                    }
                    
                    return modifiedCount;
                }
                
                // 初回実行
                var result = setupProductLinkWatcher();
                console.log("リンク変換スクリプト実行: " + result + " 個のリンクを処理しました");
                
                // インターバルIDを保存（既存のものがあれば再利用）
                if (!window._productLinkWatcherInterval) {
                    window._productLinkWatcherInterval = setInterval(function() {
                        var newCount = setupProductLinkWatcher();
                        if (newCount > 0) {
                            console.log("Found and modified " + newCount + " new product links");
                        }
                    }, 2000);
                }
                
                return result;
            """)

            print(f"リンク変換スクリプトが実行され、{result}個のリンクが処理されました")
            return result
        except Exception as e:
            print(f"リンク変換スクリプト適用中にエラー: {str(e)}")
            return 0

    @error_handler(operation="handle_point_payment_page")
    def handle_point_payment_page(self):
        """ポイント・支払い方法選択ページの処理"""
        if not self.focus_on_correct_tab():
            return False

        # 停止チェック
        if self.check_stop():
            print("ユーザーリクエストにより処理を停止します")
            self.update_status("処理を停止しました", "warning")
            return False

        print("ポイント・支払い方法選択ページの処理を開始")
        self.update_status("パスワードを入力しています", "info")

        # ポイント選択部分は完全にスキップ（デフォルトで「使用しない」が選択済み）

        # 停止チェック
        if self.check_stop():
            print("ユーザーリクエストにより処理を停止します")
            self.update_status("処理を停止しました", "warning")
            return False

        print("パスワードフィールドを待機中...")
        password_selectors = [
            "input#FormModel_Password", "input[name='FormModel.Password']",
            "input[type='password']"
        ]

        # パスワードフィールドを探して入力
        element, _ = self.wait_for_any_element(password_selectors, timeout=5)
        if not element:
            print("パスワードフィールドが見つかりませんでした。")
            self.update_status("パスワードフィールドが見つかりません", "error")
            return False

        # パスワード入力
        try:
            self.driver.execute_script(
                "arguments[0].focus(); arguments[0].value = '';", element)
            element.send_keys(self.password)
            time.sleep(0.1)
        except Exception as e:
            print(f"パスワード入力エラー: {str(e)}")
            self.update_status("パスワードの入力に失敗しました", "error")
            return False

        # 停止チェック
        if self.check_stop():
            print("ユーザーリクエストにより処理を停止します")
            self.update_status("処理を停止しました", "warning")
            return False

        print("次へボタンを待機中...")
        next_button_selectors = [
            "input[name='next'][type='image']",
            "input[type='submit'][value='次へ']", "button.next-button"
        ]

        # 次へボタンをクリック
        if not self.handle_element_action(
                next_button_selectors, "click", timeout=5, retries=3):
            self.update_status("次へボタンが見つかりませんでした", "error")
            return False

        if self.verbose_log:
            print("ページ遷移を待機中...")
        try:
            # 短いタイムアウトでの定期的な確認
            start_time = time.time()
            while time.time() - start_time < 5:
                # 停止チェック
                if self.check_stop():
                    return False

                if "payment1" in self.driver.current_url or "payment" in self.driver.current_url:
                    break
                time.sleep(0.1)
        except Exception:
            if self.verbose_log:
                print("ページ遷移が確認できませんでした")
            # 続行する（次のステップで適切に処理される）

        print("ポイント・支払い方法選択ページの処理完了")
        return True

    @error_handler(operation="handle_payment_page")
    def handle_payment_page(self):
        """支払い方法選択ページの処理（代金引換専用）"""
        if not self.focus_on_correct_tab():
            return False

        # 停止チェック
        if self.check_stop():
            print("ユーザーリクエストにより処理を停止します")
            self.update_status("処理を停止しました", "warning")
            return False

        print("支払い方法選択ページの処理を開始")
        self.update_status("代金引換を選択しています", "info")

        if self.verbose_log:
            print("代金引換ラジオボタンを待機中...")
        daibiki_selectors = [
            "input#daibiki", "input[name='daibiki']",
            "input[type='radio'][value='daibiki']"
        ]

        # 代金引換ラジオボタンを選択
        if not self.handle_element_action(
                daibiki_selectors, "click", timeout=3, retries=3):
            self.update_status("代金引換ボタンが見つかりません", "warning")
            # 既に選択されているかもしれないので、続行する

        time.sleep(0.3)

        # 停止チェック
        if self.check_stop():
            print("ユーザーリクエストにより処理を停止します")
            self.update_status("処理を停止しました", "warning")
            return False

        print("次へボタンを待機中...")
        next_button_selectors = [
            "input[name='nexttwo'][type='image']",
            "input[name='next'][type='image']",
            "input[type='submit'][value='次へ']", "button.next-button"
        ]

        # 次へボタンをクリック
        if not self.handle_element_action(
                next_button_selectors, "click", timeout=5, retries=3):
            self.update_status("次へボタンが見つかりませんでした", "error")
            return False

        self.update_status("代金引換を設定しました", "success")
        print("支払い方法選択ページの処理完了")
        return True

    @error_handler(operation="handle_recaptcha")
    def handle_recaptcha(self):
        """reCAPTCHA処理（複数商品がカートにある場合の配送方法選択も処理）"""
        if not self.focus_on_correct_tab():
            return False

        # 停止チェック
        if self.check_stop():
            print("ユーザーリクエストにより処理を停止します")
            self.update_status("処理を停止しました", "warning")
            return False

        print("配送方法とreCAPTCHA処理を開始")

        # 複数商品がカートに入っている場合の処理
        try:
            # お届け方法ラジオボタンの存在を確認
            divide_delivery_no = self.driver.find_elements(
                By.CSS_SELECTOR, "input#DivideDeliveryNo")

            if divide_delivery_no and len(divide_delivery_no) > 0:
                print("複数商品カートを検出。配送方法を選択します")
                self.update_status("複数商品の配送方法を選択中...", "info")

                # 「まとめてお届け」を選択
                self.driver.execute_script("arguments[0].click();",
                                           divide_delivery_no[0])
                print("「まとめてお届け」を選択しました")

        except Exception as e:
            print(f"配送方法選択中にエラー: {str(e)}")
            # 続行する（この機能が失敗しても他の処理は継続）
            if self.verbose_log:
                print(f"エラーの詳細: {type(e).__name__}")

        # reCAPTCHA処理
        self.update_status("reCAPTCHAの確認が必要です。チェックを入れてください。", "warning")

        if self.verbose_log:
            print("reCAPTCHAフレームを待機中...")
        try:
            # reCAPTCHAフレームを探す
            recaptcha_iframe = None
            start_time = time.time()
            while time.time() - start_time < 5:
                # 停止チェック
                if self.check_stop():
                    return False

                try:
                    recaptcha_iframe = self.driver.find_element(
                        By.CSS_SELECTOR, "iframe[title*='reCAPTCHA']")
                    if recaptcha_iframe:
                        break
                except:
                    pass
                time.sleep(0.1)

            if recaptcha_iframe:
                if self.verbose_log:
                    print("reCAPTCHAフレームに切り替え")
                self.driver.switch_to.frame(recaptcha_iframe)

                if self.verbose_log:
                    print("チェックボックスの状態を待機中...")

                # reCAPTCHA待機中も定期的に停止チェック
                max_wait_time = 60  # 最大待機時間（秒）
                wait_start_time = time.time()

                while time.time() - wait_start_time < max_wait_time:
                    # 停止チェック
                    if self.check_stop():
                        print("ユーザーリクエストにより処理を停止します")
                        self.driver.switch_to.default_content()
                        self.update_status("処理を停止しました", "warning")
                        return False

                    # reCAPTCHAの状態をチェック
                    try:
                        checkbox = self.driver.find_element(
                            By.CSS_SELECTOR, ".recaptcha-checkbox")
                        if checkbox.get_attribute("aria-checked") == "true":
                            break
                    except:
                        pass

                    # 少し待機
                    time.sleep(0.2)  # 0.5秒→0.2秒に短縮

                if self.verbose_log:
                    print("メインフレームに戻ります")
                self.driver.switch_to.default_content()
                time.sleep(0.1)
        except Exception as e:
            if self.verbose_log:
                print(f"reCAPTCHAフレーム処理エラー: {str(e)}")
            # reCAPTCHAがない場合もあるので続行する
            self.driver.switch_to.default_content()

        # 停止チェック
        if self.check_stop():
            print("ユーザーリクエストにより処理を停止します")
            self.update_status("処理を停止しました", "warning")
            return False

        print("次へボタンをクリック")
        next_button_selectors = [
            "input[name='next'][type='image']",
            "input[type='submit'][value='次へ']", "button.next-button"
        ]

        # 次へボタンをクリック
        if not self.handle_element_action(
                next_button_selectors, "click", timeout=5, retries=3):
            self.update_status("次へボタンが見つかりませんでした", "error")
            return False

        print("配送方法とreCAPTCHA処理完了")
        return True

    @error_handler(operation="start_automation")
    def start_automation(self):
        """現在のページから自動化を開始"""
        # 購入処理中フラグを設定
        self.purchase_in_progress = True

        # 現在のタブに焦点を合わせる (変更なし)
        if hasattr(self, 'product_tab'):
            self.driver.switch_to.window(self.product_tab)
        else:
            if not self.focus_on_correct_tab():
                self.update_status("マップカメラのタブが見つかりません。マップカメラサイトを開いてください。",
                                   "error")
                return False

        print("自動化を開始します")
        current_url = self.driver.current_url
        print(f"現在のURL: {current_url}")

        # 停止チェック (変更なし)
        if self.check_stop():
            print("ユーザーリクエストにより処理を停止します")
            self.update_status("処理を停止しました", "warning")
            return False

        # 商品詳細ページかどうかを確認 (変更なし)
        if self.is_product_page(
                current_url) or self.is_product_page_by_content():
            print("商品詳細ページから自動化を開始します")
            self.update_status("自動購入処理を開始します...", "info")
        else:
            print("対応していないページです")
            self.update_status("対応していないページです。商品詳細ページで実行してください。", "error")
            return False

        # *** SOLD OUT検出処理 *** (変更なし)
        print("SOLD OUTチェックを開始")
        if self.is_sold_out():
            print("商品はSOLD OUTです")
            self.update_status("この商品はSOLD OUTです。", "warning")
            return False

        # カートに追加 (変更なし)
        print("カートに追加処理を開始")
        if self.check_stop():
            print("ユーザーリクエストにより処理を停止します")
            self.update_status("処理を停止しました", "warning")
            return False

        self.update_status("カートに商品を追加します", "info")
        cart_button_selectors = [
            "input[name='cartPut']", "button.cart-button", "a.add-to-cart"
        ]

        if not self.handle_element_action(
                cart_button_selectors, "click", timeout=5, retries=3):
            self.update_status("カートボタンが見つかりませんでした", "error")
            return False
        time.sleep(0.1)

        if self.check_stop():
            print("ユーザーリクエストにより処理を停止します")
            self.update_status("処理を停止しました", "warning")
            return False

        # ここから変更: レジに進むボタンクリックの代わりに直接URLに遷移
        print("裏技: お届け先設定画面をスキップして直接ポイント・支払い方法選択画面へ移動します")
        self.update_status("高速モード: お届け先設定画面をスキップします", "info")

        try:
            # 直接ポイントと支払い方法選択画面へ移動
            self.driver.get(
                "https://www.mapcamera.com/ec/cart/order/pointandpayment")

            # ページが完全に読み込まれるのを待機
            WebDriverWait(self.driver, 5).until(
                lambda d: d.execute_script(
                    "return document.readyState") == "complete"
            )

            print("ポイント・支払い方法選択画面に直接移動しました")
            self.update_status("ポイント・支払い方法選択画面に移動しました", "success")
        except Exception as e:
            self.log_error("ポイント・支払い方法選択画面への直接移動でエラー",
                           e, operation="direct_to_payment")
            self.update_status("高速移動に失敗しました。通常モードで続行します。", "warning")

            # 失敗した場合は通常のフローでレジに進む
            print("通常モード: レジに進む処理を開始")
            self.update_status("レジに進みます", "info")
            checkout_button_selectors = [
                "a#checkout2", "a.checkout-button", "a[href*='checkout']",
                "button.proceed-to-checkout"
            ]

            if not self.handle_element_action(
                    checkout_button_selectors, "click", timeout=5, retries=3):
                self.update_status("レジへ進むボタンが見つかりませんでした", "error")
                return False
            time.sleep(0.1)

            # 配送情報ページでreCAPTCHA
            print("現在のURL確認: delivery")
            if "/delivery" in self.driver.current_url:
                if not self.handle_recaptcha():
                    return False

        if self.check_stop():
            print("ユーザーリクエストにより処理を停止します")
            self.update_status("処理を停止しました", "warning")
            return False

        time.sleep(0.1)

        # ポイント・支払い方法選択ページ (変更なし)
        print("現在のURL確認: pointandpayment")
        if "/pointandpayment" in self.driver.current_url:
            if not self.handle_point_payment_page():
                return False

        if self.check_stop():
            print("ユーザーリクエストにより処理を停止します")
            self.update_status("処理を停止しました", "warning")
            return False

        time.sleep(0.1)

        # 支払い方法選択 (変更なし)
        print("現在の URL確認: payment1")
        if "/payment1" in self.driver.current_url or "/payment" in self.driver.current_url:
            if not self.handle_payment_page():
                return False

        # 自動化完了 (変更なし)
        print("自動化完了")
        self.update_status(
            "処理が完了しました。注文を確定する場合は画面の「注文を確定する」ボタンをクリックし、次の商品を選択する場合は商品一覧タブに切り替えてください。", "success")

        # 購入処理中フラグを維持（最終確認画面での誤メッセージ防止）
        # self.purchase_in_progress = True のまま保持

        # タブ切り替え防止フラグを設定 (変更なし)
        self.prevent_tab_switch = True

        # 監視関連のフラグをリセット (変更なし)
        if hasattr(self, 'is_monitoring'):
            self.is_monitoring = False

        return True

    def cleanup(self):
        """ブラウザを終了 - 改良版（セッション終了の適切な検出）"""
        try:
            print("クリーンアップを開始します")
            if hasattr(self, 'driver'):
                try:
                    # セッションが有効かどうかを最初に確認
                    session_active = False
                    try:
                        # 簡単なコマンドを実行して接続をテスト
                        self.driver.current_window_handle
                        session_active = True
                    except Exception:
                        # 例外が発生したらセッションは既に終了している
                        session_active = False
                        print("WebDriverセッションは既に終了しています")

                    # セッションがアクティブな場合のみメッセージ消去を試みる
                    if session_active:
                        try:
                            self.driver.execute_script("""
                                var msg = document.getElementById('automation-message');
                                if (msg) msg.remove();
                            """)
                        except Exception as e:
                            print(f"メッセージ消去中にエラーが発生しましたが処理を継続します: {str(e)}")
                except Exception as e:
                    print(f"WebDriver状態確認中にエラー: {str(e)}")
                    # エラーが発生しても処理を続行

                print("クリーンアップが完了しました")
        except Exception as e:
            print(f"クリーンアップ処理全体でエラー: {str(e)}")
            if self.verbose_log:
                print(f"エラーの詳細: {type(e).__name__}")

    def tab_exists(self, tab_handle):
        """タブが存在するかどうかを確認"""
        try:
            return tab_handle in self.driver.window_handles
        except:
            return False

    def safe_switch_to_tab(self, tab_handle):
        """安全にタブを切り替える（タブが存在する場合のみ）"""
        if self.tab_exists(tab_handle):
            self.driver.switch_to.window(tab_handle)
            return True
        return False

    def monitor_page_updates(self, url=None, callback=None):
        """
        指定したURLのページを監視し、更新があれば通知する
        既存のマップカメラタブを利用するよう変更

        Args:
            url (str, optional): 監視するURL。Noneの場合は既存のタブを使用
            callback (callable): 更新検出時に呼び出すコールバック関数
        """
        try:
            # 現在のウィンドウハンドルを保存
            current_handle = self.driver.current_window_handle
            monitor_tab = None

            # まず既存のマップカメラタブを探す
            for handle in self.driver.window_handles:
                try:
                    self.driver.switch_to.window(handle)
                    if "mapcamera.com" in self.driver.current_url:
                        # 既存のマップカメラタブを発見
                        if self.is_product_list_page(self.driver.current_url):
                            # 検索結果/商品一覧ページなら、このタブを使用
                            monitor_tab = handle
                            print(
                                f"既存の商品一覧タブを監視に使用します: {self.driver.current_url}"
                            )
                            # URLが指定されていても、そこに移動しない（既存タブの内容を尊重）
                            break
                except:
                    continue

            # 既存のタブが見つからず、URLが指定されている場合は新規タブを開く
            if not monitor_tab and url:
                print("既存の商品一覧タブが見つからないため、新しいタブで開きます")
                self.driver.execute_script(f"window.open('{url}', '_blank');")
                # 新しいタブに切り替え
                handles = self.driver.window_handles
                monitor_tab = handles[-1]
                self.driver.switch_to.window(monitor_tab)

            if not monitor_tab:
                print("監視するタブが見つかりませんでした")
                self.update_status("監視するタブが見つかりません。マップカメラの商品一覧ページを開いてください。",
                                   "error")
                return False

            # ページが完全に読み込まれるまで待機
            WebDriverWait(self.driver, 10).until(lambda d: d.execute_script(
                "return document.readyState") == "complete")

            # 初期状態の商品リスト情報を保存
            initial_products = self._get_product_list_info()

            # 設定から監視間隔を取得
            monitoring_interval = self.config.get("monitoring_interval", 10)

            print(f"ページ監視を開始しました（間隔: {monitoring_interval}秒）")
            self.update_status(
                f"商品更新の監視を開始しました（{monitoring_interval}秒間隔）。更新を検出したらお知らせします。",
                "info")

            # 監視情報を保存
            self.monitor_tab = monitor_tab
            self.monitor_callback = callback
            self.stop_requested = False
            self.last_check_result = initial_products

            # 元のタブに戻る
            self.driver.switch_to.window(current_handle)

            # 監視スレッドを開始
            self.monitor_thread = threading.Thread(target=self._monitor_loop,
                                                   daemon=True)
            self.monitor_thread.start()

            return True

        except Exception as e:
            self.log_error("ページ監視の開始に失敗しました",
                           e,
                           operation="monitor_page_updates")
            # 元のタブに戻る
            try:
                self.driver.switch_to.window(current_handle)
            except:
                pass
            return False

    def _monitor_loop(self):
        """バックグラウンドで監視を実行するループ - タブ切り替え問題の修正"""
        monitoring_interval = self.config.get(
            "monitoring_interval", 10)  # ここを変更
        consecutive_errors = 0   # 連続エラー回数
        first_successful_check = False  # 最初の正常な検出フラグ
        update_detected = False  # 更新検出フラグを追加

        while not self.stop_requested and hasattr(self, 'monitor_tab'):
            cycle_start_time = time.time()  # サイクル開始時間を記録
            try:
                # 現在のタブを保存
                try:
                    current_handle = self.driver.current_window_handle
                except:
                    # 現在のタブが取得できない場合、監視タブを現在のタブとする
                    if self.tab_exists(self.monitor_tab):
                        current_handle = self.monitor_tab
                    else:
                        # 監視タブも存在しない場合は終了
                        print("現在のタブと監視タブの両方が見つかりません。監視を終了します。")
                        break

                # ここから新しく追加するコード ↓
                # タブ切り替え防止フラグをチェック
                if hasattr(self, 'prevent_tab_switch') and self.prevent_tab_switch:
                    if self.verbose_log:
                        print("タブ切り替え防止フラグが有効なため、モニタリングタブへの切り替えをスキップします")
                    # タブ切り替えせずに次のサイクルへ
                    time.sleep(monitoring_interval)
                    continue
                # ここまでが新しく追加するコード ↑

                # 監視タブが存在するか確認
                if not self.tab_exists(self.monitor_tab):
                    print("監視タブが閉じられました。監視を終了します。")
                    break

                # ===== タブ切り替え問題対応 =====
                # 現在のタブが監視タブでない場合のみ切り替える
                same_tab = (current_handle == self.monitor_tab)
                if not same_tab:
                    # 監視タブに切り替え
                    self.driver.switch_to.window(self.monitor_tab)

                # リロードの前に一時的なフラグを保存
                temp_js_var = f"window.__monitoring_check_{int(time.time())}"
                self.driver.execute_script(f"{temp_js_var} = true;")

                # ページをリロード
                self.driver.refresh()

                # ページ読み込み待機
                try:
                    WebDriverWait(
                        self.driver, 10).until(lambda d: d.execute_script(
                            "return document.readyState") == "complete")
                    time.sleep(1)  # 追加の待機
                except Exception as e:
                    print(f"ページ読み込み待機でタイムアウト: {str(e)}")
                    time.sleep(2)

                # 更新の検出処理
                # ページの商品情報を取得
                current_data = self._get_product_list_info()

                # 前回の結果と比較（初回は比較しない）
                if hasattr(self,
                           'last_check_result') and self.last_check_result:
                    # 商品情報の比較ロジックを改善（一例）
                    changes_detected = self._detect_product_changes(
                        self.last_check_result, current_data)

                    if changes_detected:
                        print("商品の更新を検出しました！")

                        # コールバック関数の呼び出し
                        if hasattr(self, 'monitor_callback') and callable(
                                self.monitor_callback):
                            self.monitor_callback()

                        # 更新検出フラグを設定
                        update_detected = True

                        # 商品は1日1回しか更新されないため、監視を自動停止
                        print("商品更新が検出されたため、監視を自動停止します")
                        self.stop_requested = True
                        break

                # 結果を記録
                self.last_check_result = current_data

                # 元のタブに戻る（監視タブと異なる場合のみ）
                if not same_tab and self.tab_exists(current_handle):
                    self.driver.switch_to.window(current_handle)

                # エラーリセット
                consecutive_errors = 0

                # 処理にかかった時間を計算
                elapsed_time = time.time() - cycle_start_time

                # 残りの待機時間を計算（最小0.5秒を保証）
                remaining_wait = max(0.5, monitoring_interval - elapsed_time)

                if self.verbose_log:
                    print(
                        f"監視処理時間: {elapsed_time:.2f}秒、残り待機時間: {remaining_wait:.2f}秒"
                    )

                # 適切な時間だけ待機
                time.sleep(remaining_wait)

            except Exception as e:
                # エラー処理
                self.log_error("監視ループでエラーが発生", e, operation="_monitor_loop")
                consecutive_errors += 1

                try:
                    # エラー発生時も元のタブに戻る（監視タブと異なる場合のみ）
                    if not same_tab and current_handle and self.tab_exists(
                            current_handle):
                        self.driver.switch_to.window(current_handle)
                except:
                    pass

                # 連続エラー時は待機時間を延長
                error_wait = min(5, consecutive_errors)
                time.sleep(error_wait)

        # 監視終了時の処理を追加
        if update_detected:
            # 1日の更新が完了したことを通知
            print("本日の商品更新は検出されました。監視を終了します。")

    def _get_product_list_info(self):
        """商品リストの情報を取得"""
        try:
            # MapCameraサイトの構造に基づいてDOM操作
            product_data = self.driver.execute_script("""
                var result = {
                    count: 0,
                    items: []
                };
                
                // 商品リストコンテナを取得
                var container = document.querySelector('ul.srcitemlist');
                if (!container) return result;
                
                // 商品アイテムを取得
                var items = container.querySelectorAll('li.item_wrap');
                result.count = items.length;
                
                // 最初の10件のみ処理（パフォーマンス向上）
                for (var i = 0; i < Math.min(10, items.length); i++) {
                    var item = items[i];
                    
                    // 商品ID
                    var id = item.getAttribute('data-mapcode') || '';
                    
                    // 商品名
                    var nameElem = item.querySelector('.txt > a');
                    var name = nameElem ? nameElem.textContent.trim() : '';
                    
                    // 価格情報
                    var priceElem = item.querySelector('.price > span > span > b');
                    var price = priceElem ? priceElem.textContent.trim() : '';
                    
                    // SOLD OUT状態
                    var soldOutElem = item.querySelector('.price');
                    var isSoldOut = soldOutElem ? soldOutElem.textContent.includes('SOLD OUT') : false;
                    
                    // 商品リンク
                    var linkElem = item.querySelector('.itembox > a');
                    var link = linkElem ? linkElem.getAttribute('href') : '';
                    
                    result.items.push({
                        id: id,
                        name: name,
                        price: price,
                        soldOut: isSoldOut,
                        link: link
                    });
                }
                
                return result;
            """)

            # タイムスタンプを追加
            product_data['timestamp'] = time.time()

            return product_data
        except Exception as e:
            self.log_error("商品リスト取得エラー", e, operation="_get_product_list_info")
            return {"count": 0, "items": [], "timestamp": time.time()}

    def _detect_product_changes(self, previous_data, current_data):
        """前回と今回の商品リスト情報を比較して変更を検出する"""
        try:
            # 基本的な比較（商品数の変化）
            if previous_data.get('count', 0) != current_data.get('count', 0):
                if self.verbose_log:
                    print(
                        f"商品数の変化を検出: {previous_data.get('count', 0)} -> {current_data.get('count', 0)}"
                    )
                return True

            # タイムスタンプ比較（あまりに短時間での再チェックは無視）
            if ('timestamp' in previous_data and 'timestamp' in current_data
                    and current_data['timestamp'] - previous_data['timestamp']
                    < 5):
                return False

            # 商品アイテムの比較（最初の数個のみ）
            prev_items = previous_data.get('items', [])
            curr_items = current_data.get('items', [])

            # 商品が1つもない場合はスキップ
            if not prev_items or not curr_items:
                return False

            # 商品IDのセットを比較
            prev_ids = {
                item.get('id', '')
                for item in prev_items if item.get('id')
            }
            curr_ids = {
                item.get('id', '')
                for item in curr_items if item.get('id')
            }

            # 新しい商品が追加されたかチェック
            new_ids = curr_ids - prev_ids
            if new_ids:
                if self.verbose_log:
                    print(f"新しい商品IDを検出: {new_ids}")
                return True

            # 商品の順序が変わったかチェック
            for i in range(min(len(prev_items), len(curr_items),
                               3)):  # 最初の3つをチェック
                if (prev_items[i].get('id') != curr_items[i].get('id')
                        or prev_items[i].get('soldOut')
                        != curr_items[i].get('soldOut')):
                    if self.verbose_log:
                        print(f"商品順序またはSOLD OUT状態の変化を検出")
                    return True

            return False
        except Exception as e:
            print(f"商品変更検出でエラー: {str(e)}")
            # エラーの場合は安全側に倒して変更なしと判断
            return False

    def set_password(self, password):
        """パスワードを設定する"""
        self.password = password
        return True

    def _get_product_list_info(self):
        """商品リストの情報を取得"""
        try:
            # MapCameraサイトの構造に基づいてDOM操作
            product_data = self.driver.execute_script("""
                var result = {
                    count: 0,
                    items: []
                };
                
                // 商品リストコンテナを取得
                var container = document.querySelector('ul.srcitemlist');
                if (!container) return result;
                
                // 商品アイテムを取得
                var items = container.querySelectorAll('li.item_wrap');
                result.count = items.length;
                
                // 最初の10件のみ処理（パフォーマンス向上）
                for (var i = 0; i < Math.min(10, items.length); i++) {
                    var item = items[i];
                    
                    // 商品ID
                    var id = item.getAttribute('data-mapcode') || '';
                    
                    // 商品名
                    var nameElem = item.querySelector('.txt > a');
                    var name = nameElem ? nameElem.textContent.trim() : '';
                    
                    // 価格情報
                    var priceElem = item.querySelector('.price > span > span > b');
                    var price = priceElem ? priceElem.textContent.trim() : '';
                    
                    // SOLD OUT状態
                    var soldOutElem = item.querySelector('.price');
                    var isSoldOut = soldOutElem ? soldOutElem.textContent.includes('SOLD OUT') : false;
                    
                    // 商品リンク
                    var linkElem = item.querySelector('.itembox > a');
                    var link = linkElem ? linkElem.getAttribute('href') : '';
                    
                    result.items.push({
                        id: id,
                        name: name,
                        price: price,
                        soldOut: isSoldOut,
                        link: link
                    });
                }
                
                return result;
            """)

            # タイムスタンプを追加
            product_data['timestamp'] = time.time()

            return product_data
        except Exception as e:
            self.log_error("商品リスト取得エラー", e, operation="_get_product_list_info")
            return {"count": 0, "items": [], "timestamp": time.time()}

    def stop_monitoring(self, async_mode=True):
        """ページ監視を停止する（強化版）"""
        try:
            # まず停止フラグを設定
            self.stop_requested = True

            # 監視タブの参照を保存（後でクリーンアップするため）
            monitor_tab_ref = getattr(self, 'monitor_tab', None)

            # 非同期モードの処理
            if async_mode:
                if self.verbose_log:
                    print("監視停止をリクエストしました（バックグラウンドで処理中）")
                else:
                    print("監視停止処理を開始しました")

                # すべての監視関連変数を明示的にリセット
                self._reset_monitoring_state()
                return True
        except Exception as e:
            self.log_error("監視停止でエラーが発生", e, operation="stop_monitoring")
            return False

    def _reset_monitoring_state(self):
        """監視関連の状態を完全にリセットする"""
        # 監視関連の属性をクリア
        for attr in ['monitor_tab', 'monitor_url', 'last_check_result',
                     'monitor_callback', 'monitor_thread', 'is_monitoring']:
            if hasattr(self, attr):
                delattr(self, attr)

        # タブ参照の競合を防ぐため、list_tabとproduct_tabもリセット
        # これらは連続購入モードで再設定される
        for tab_attr in ['list_tab', 'product_tab']:
            if hasattr(self, tab_attr):
                delattr(self, tab_attr)

    def _cleanup_monitoring_resources(self):
        """監視リソースをクリーンアップする（拡張版）"""
        # 監視関連の属性をクリア（既存の処理）
        for attr in ['monitor_tab', 'monitor_url', 'last_check_result', 'monitor_callback', 'monitor_thread']:
            if hasattr(self, attr):
                delattr(self, attr)

        # 追加: タブ参照の競合を防ぐため、連続購入モードで使用する可能性のあるタブ参照もリセット
        for tab_attr in ['list_tab', 'product_tab']:
            if hasattr(self, tab_attr):
                # リセットせずに記録だけする場合は以下をコメント解除
                # self.log(f"タブ参照 {tab_attr} を保持: {getattr(self, tab_attr)}")
                delattr(self, tab_attr)

        # 既存の処理を続行
        if self.verbose_log:
            print("監視リソースをクリーンアップしました")

        if hasattr(self, 'update_status'):
            self.update_status("ページ監視を停止しました", "info")

            print("ページ監視を停止しました")
            if hasattr(self, 'update_status'):
                self.update_status("ページ監視を停止しました", "info")


# 直接実行された場合の処理
if __name__ == "__main__":
    print("このスクリプトは直接実行せず、GUIから利用してください。")
    print("GUIを起動するには、mapcamera_gui.pyを実行してください。")
