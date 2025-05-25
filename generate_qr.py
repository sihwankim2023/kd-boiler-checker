import qrcode

# QR 코드에 넣을 URL
url = "https://kd-boiler-checker-63jr3mw5k8dxjkj22mvtzy.streamlit.app/"

# QR 코드 생성
qr = qrcode.QRCode(
    version=1,
    error_correction=qrcode.constants.ERROR_CORRECT_L,
    box_size=10,
    border=4,
)
qr.add_data(url)
qr.make(fit=True)

# QR 코드 이미지 생성
qr_image = qr.make_image(fill_color="black", back_color="white")

# 이미지 저장
qr_image.save("kd-boiler-qr.png")
print("QR 코드가 생성되었습니다. 'kd-boiler-qr.png' 파일을 확인해주세요.") 