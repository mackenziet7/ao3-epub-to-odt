import qrcode

def generate_qr_png(url, out_path):
    """Write a QR PNG for `url` to `out_path`. Raises on failure."""
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=2,
    )
    qr.add_data(url)
    qr.make(fit=True)
    qr.make_image(fill_color="black", back_color="white").save(str(out_path))