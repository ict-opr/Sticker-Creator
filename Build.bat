pyinstaller --onefile --windowed ^
  --icon=Icons\Icon.ico ^
  --add-data "Fonts;Fonts" ^
  --collect-submodules reportlab.graphics.barcode ^
  --collect-data reportlab ^
  --name "Sticker Creator" "pyStickerCreator.py"
