import logging
import mss
import numpy as np
import pytesseract
import cv2
import time
import tkinter as tk
from PIL import Image, ImageEnhance, ImageOps
import json
import sys

# ロガーの設定
logger = logging.getLogger(__name__)  # このモジュール専用のロガー
logger.setLevel(logging.DEBUG)

# ログ出力先の設定
file_handler = logging.FileHandler("ocr_system.log", mode='w', encoding='utf-8')
file_handler.setLevel(logging.DEBUG)

console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)

# ログフォーマットの設定
formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

# ロガーにハンドラを追加
logger.addHandler(file_handler)
logger.addHandler(console_handler)

def capture_screen():
    """
    スクリーン全体をキャプチャし、画像として返す。
    - スクリーンショットを2倍に拡大して詳細な解析が可能にする。
    """
    try:
        with mss.mss() as sct:
            monitor = sct.monitors[0]  # メインモニタ全体
            screenshot = sct.grab(monitor)
            img = np.array(screenshot)

        # スクリーンショットを2倍に拡大
        img = cv2.resize(img, (img.shape[1] * 2, img.shape[0] * 2), interpolation=cv2.INTER_LINEAR)
        logger.debug("Screen captured and resized successfully.")
        return img
    except Exception as e:
        logger.error(f"Error capturing screen: {e}")
        return None

def preprocess_image(image):
    """
    画像の前処理を行い、OCR解析に適した形式に変換する。
    - ステップ:
        1. グレースケール化
        2. ノイズ除去 (GaussianBlur)
        3. コントラスト強調 (Pillowを使用)
        4. 色反転
        5. 二値化 (OTSU法)
    """
    try:
        # グレースケール化
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        # ノイズ除去
        blur = cv2.GaussianBlur(gray, (5, 5), 0)
        # コントラスト強調 (PILを利用)
        pil_image = Image.fromarray(blur)
        enhancer = ImageEnhance.Contrast(pil_image)
        enhanced_image = enhancer.enhance(2.0)  # コントラストを強調（2.0は強調率）
        # 色反転
        inverted_image = ImageOps.invert(enhanced_image)
        # OpenCV形式に戻す
        processed_image = np.array(inverted_image)
        # 二値化 (OTSU)
        _, binary = cv2.threshold(processed_image, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        logger.debug("Image preprocessing completed.")
        return binary
    except Exception as e:
        logger.error(f"Error during image preprocessing: {e}")
        return None

""" def correct_text_spacing_with_cleaning(text):
    ""
    不要なスペースや記号を除去し、単語を結合する
    ""
    try:
        # ノイズを削除（記号や複数スペース）
        cleaned_text = re.sub(r'[^\w\s\u4e00-\u9fff]', '', text)  # 日本語と英数字以外を削除
        cleaned_text = re.sub(r'\s+', '', cleaned_text)  # 複数のスペースを削除
        logger.debug(f"Corrected and cleaned text: {cleaned_text}")
        return cleaned_text
    except Exception as e:
        logger.error(f"Error correcting text spacing: {e}")
        return text """


def detect_text(image):
    """
    OCRでテキストを検出し、グループ化して返す。
    - Tesseractを使用してテキストデータを解析。
    - 結果のバウンディングボックスを元画像スケールに戻す。
    - テキストをグループ化し、解析結果を整理。
    """
    try:
        custom_config = r'--oem 3 --psm 11 -c preserve_interword_spaces=1'
        text_data = pytesseract.image_to_data(image, lang='jpn', config=custom_config, output_type=pytesseract.Output.DICT)

        # バウンディングボックスを元画像サイズにスケーリング（1/2）
        scale_factor = 2
        for i in range(len(text_data['left'])):
            text_data['left'][i] //= scale_factor
            text_data['top'][i] //= scale_factor
            text_data['width'][i] //= scale_factor
            text_data['height'][i] //= scale_factor

        # 文字をグループ化
        grouped_text, grouped_indices, grouped_boxes = group_text_blocks(text_data)
        logger.debug(f"OCR grouped text: {grouped_text}")

        return grouped_text, grouped_indices, grouped_boxes, text_data
    except Exception as e:
        logger.error(f"Error during OCR detection: {e}")
        return None, None, None, None

def group_text_blocks(text_data, distance_threshold=20):
    """
    OCR解析結果をもとに文字列をグループ化し、それぞれのバウンディングボックスを計算する。
    - グループ化は文字列間の距離に基づいて行う。
    - 各グループのバウンディングボックスを結合して計算。

    Args:
        text_data (dict): Tesseractの出力データ。
        distance_threshold (int): グループ化のための距離閾値。

    Returns:
        tuple: グループ化されたテキスト、インデックス、バウンディングボックス。
    """
    try:
        n = len(text_data['text'])
        grouped_text = []
        grouped_indices = []
        grouped_boxes = []  # グループ全体のバウンディングボックス
        current_group = ""
        current_indices = []
        current_box = None  # 現在のグループのバウンディングボックス
        prev_right = 0

        for i in range(n):
            word = text_data['text'][i].strip()
            if not word:
                continue

            x = text_data['left'][i]
            y = text_data['top'][i]
            width = text_data['width'][i]
            height = text_data['height'][i]
            right = x + width

            # 新しいバウンディングボックスを開始
            if current_group and (x - prev_right) >= distance_threshold:
                grouped_text.append(current_group)
                grouped_indices.append(current_indices)
                grouped_boxes.append(current_box)

                current_group = word
                current_indices = [i]
                current_box = (x, y, width, height)
            else:
                # 既存のグループに追加
                current_group += word
                current_indices.append(i)

                # バウンディングボックスを更新
                if current_box:
                    min_x = min(current_box[0], x)
                    min_y = min(current_box[1], y)
                    max_x = max(current_box[0] + current_box[2], right)
                    max_y = max(current_box[1] + current_box[3], y + height)
                    current_box = (min_x, min_y, max_x - min_x, max_y - min_y)
                else:
                    current_box = (x, y, width, height)

            prev_right = right

        # 最後のグループを保存
        if current_group:
            grouped_text.append(current_group)
            grouped_indices.append(current_indices)
            grouped_boxes.append(current_box)

        logger.debug(f"Grouped text: {grouped_text} with boxes: {grouped_boxes}")
        return grouped_text, grouped_indices, grouped_boxes
    except Exception as e:
        logger.error(f"Error in grouping text blocks: {e}")
        return [], [], []

def search_text_in_blocks(grouped_text, grouped_boxes, target_strings):
    """
    グループ化されたテキストを検索し、ターゲット文字列に一致するバウンディングボックスを取得する。

    Args:
        grouped_text (list): グループ化されたテキストリスト。
        grouped_boxes (list): 各グループのバウンディングボックスリスト。
        target_strings (list): 検索対象のターゲット文字列リスト。

    Returns:
        list: 一致するバウンディングボックスのリスト。
    """
    try:
        boxes = []
        for i, word in enumerate(grouped_text):
            for target in target_strings:
                if target in word:
                    logger.info(f"Match found: '{word}' contains target '{target}'")
                    boxes.append(grouped_boxes[i])
                    break

        if not boxes:
            logger.warning(f"No matches found for target strings: {target_strings}")
        return boxes
    except Exception as e:
        logger.error(f"Error during grouped text search: {e}")
        return []

def highlight_text(canvas, boxes):
    """
    Tkinter Canvas上にバウンディングボックスを描画してテキストをハイライト表示する。

    Args:
        canvas (tk.Canvas): 描画対象のTkinter Canvasオブジェクト。
        boxes (list): ハイライトするバウンディングボックスリスト。
    """
    canvas.delete("highlight")  # 前回の描画を削除
    for box in boxes:
        x, y, w, h = box
        canvas.create_rectangle(x, y, x + w, y + h, outline="red", width=3, tags="highlight")
    logger.debug(f"Highlighted {len(boxes)} text boxes on canvas.")

def initialize_display_window():
    """
    OCR結果を表示するためのTkinterウィンドウを初期化。

    Returns:
        tuple: TkinterのルートウィンドウとCanvasオブジェクト。
    """
    root = tk.Tk()
    root.attributes('-fullscreen', True)
    root.attributes('-topmost', True)
    root.attributes('-transparentcolor', 'white')  # 背景を透明にする
    root.configure(bg='white')  # 背景色を透明色に設定
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    canvas = tk.Canvas(root, width=screen_width, height=screen_height, bg="white", highlightthickness=0)
    canvas.pack(fill=tk.BOTH, expand=True)
    return root, canvas

def visualize_bounding_boxes(image, text_data):
    """
    OCR結果のバウンディングボックスを画像に描画して保存する。

    Args:
        image (numpy.ndarray): 入力画像。
        text_data (dict): Tesseractの出力データ。
    """
    try:
        # バウンディングボックスを描画
        for i in range(len(text_data['text'])):

            scale_factor = 2
            text_data['left'][i] *= scale_factor
            text_data['top'][i] *= scale_factor
            text_data['width'][i] *= scale_factor
            text_data['height'][i] *= scale_factor

            x, y, w, h = text_data['left'][i], text_data['top'][i], text_data['width'][i], text_data['height'][i]
            if text_data['text'][i].strip():  # 空文字を除外
                cv2.rectangle(image, (x, y), (x + w, y + h), (0, 255, 0), 2)  # 緑色で枠を描画

        # 画像を保存
        output_path = "output_with_bounding_boxes.png"
        cv2.imwrite(output_path, image)
        logger.info(f"Bounding boxes visualized and saved to {output_path}")
    except Exception as e:
        logger.error(f"Error visualizing bounding boxes: {e}")


def main():
        # コマンドライン引数でターゲット文字列リストを受け取る
    if len(sys.argv) > 1:
        try:
            # JSON形式の文字列をリストに復元
            target_strings = json.loads(sys.argv[1])
            logger.info(f"Received target strings: {target_strings}")
        except json.JSONDecodeError as e:
            logger.error(f"Failed to decode target strings: {e}")
            return
    else:
        logger.error("No target strings provided. Exiting...")
        return

    # ウィンドウ初期化
    root, canvas = initialize_display_window()

    # 終了フラグ
    terminate_event = tk.BooleanVar(master=root, value=False)
    root.bind("<KeyPress-q>", lambda event: terminate_event.set(True))  # qキーで終了

    last_time = time.time()

    def update_frame():
        nonlocal last_time

        # スクリーンキャプチャと前処理
        screen_image = capture_screen()
        if screen_image is None:
            logger.error("Failed to capture screen. Skipping frame update.")
            root.after(10, update_frame)
            return

        processed_image = preprocess_image(screen_image)
        if processed_image is None:
            logger.error("Failed to preprocess image. Skipping frame update.")
            root.after(10, update_frame)
            return

        # OCRを1秒ごとに実行
        if time.time() - last_time > 1:
            grouped_text, grouped_indices, grouped_boxes, text_data = detect_text(processed_image)
            last_time = time.time()

            if grouped_text is not None and grouped_boxes is not None and text_data is not None:
                # テキスト検索とハイライト
                boxes = search_text_in_blocks(grouped_text, grouped_boxes, target_strings)
                highlight_text(canvas, boxes)
                # バウンディングボックスの可視化と保存
                visualize_bounding_boxes(screen_image, text_data)

        if terminate_event.get():
            root.destroy()
            logger.info("System terminated by user.")
        else:
            root.after(10, update_frame)

    logger.info("System started.")
    update_frame()
    root.mainloop()

if __name__ == '__main__':
    main()