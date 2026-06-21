from flask import Flask, request, send_file, jsonify
from PIL import Image
import requests
from io import BytesIO
import os

app = Flask(__name__)

# إعدادات القص (نفس القيم من draw_xy.py)
PAD_LEFT_PCT = 0.40
PAD_RIGHT_PCT = 0.11
PAD_TOP_PCT = 0.17
PAD_BOTTOM_PCT = 0.05

@app.route('/crop', methods=['GET'])
def crop_image():
    # قراءة المعاملات من الطلب
    image_url = request.args.get('image_url')
    x = request.args.get('x', type=float)
    y = request.args.get('y', type=float)
    w = request.args.get('w', type=float)    # العرض
    h = request.args.get('h', type=float)    # الارتفاع

    if not image_url or x is None or y is None or w is None or h is None:
        return jsonify({'error': 'Missing parameters'}), 400

    try:
        # تحميل الصورة من الرابط
        response = requests.get(image_url, timeout=10)
        response.raise_for_status()
        img = Image.open(BytesIO(response.content))
        img_w, img_h = img.size

        # تحويل الإحداثيات إلى بكسلات (إذا كانت نسب مئوية)
        left_px = x * img_w if x <= 1.0 else x
        top_px = y * img_h if y <= 1.0 else y
        box_w_px = w * img_w if w <= 1.0 else w
        box_h_px = h * img_h if h <= 1.0 else h

        # تطبيق الهوامش (نفس منطق draw_xy.py)
        final_left = max(0, left_px - (img_w * PAD_LEFT_PCT))
        final_right = min(img_w, left_px + box_w_px + (img_w * PAD_RIGHT_PCT))
        final_top = max(0, top_px - (img_h * PAD_TOP_PCT))
        final_bottom = min(img_h, top_px + box_h_px + (img_h * PAD_BOTTOM_PCT))

        if final_right <= final_left or final_bottom <= final_top:
            return jsonify({'error': 'Invalid crop region'}), 400

        crop = img.crop((final_left, final_top, final_right, final_bottom))
        if crop.mode in ("RGBA", "P"):
            crop = crop.convert("RGB")

        # حفظ الصورة في BytesIO وإرجاعها
        img_io = BytesIO()
        crop.save(img_io, 'JPEG', quality=85)
        img_io.seek(0)
        return send_file(img_io, mimetype='image/jpeg')

    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)