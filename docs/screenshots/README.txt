このフォルダには、test-fixtures/pptx-export.md を使って生成した
HTML/PPTX スライドの比較スクリーンショットを配置します。

生成方法:
  1. npm run build でビルド
  2. node src/native-pptx/tools/gen-pptx.js <marpmd> <outputdir> で PPTX 生成
  3. node src/native-pptx/tools/compare-visuals.js <html> <pptx> <chrome> で比較画像生成
  4. 生成された compare-NNN.png, html-slide-NNN.png, pptx-slide-NNN.png をここに配置
