# 再生成用ソース

第1回 ミラキン業界研究会「金属業界」全41枚のPowerPointを生成するコードです。

```bash
cd output/source
npm install
npm run build        # assets/*.png を描画 → ../MIRAKIN_金属業界研究会_全41枚.pptx を出力
```

| ファイル | 役割 |
|---|---|
| `content.js` | 章構成、スライドタイトル、発表者ノート（講師原稿・想定回答・注意・出典）。文言の修正は主にここ |
| `build.js` | 41枚のレイアウト。テーマ色・フォント・スライドマスター（LIGHT / WORK / DARK / COVER / HERO_GD / HERO_END）もここで定義 |
| `scenes.js` | フラットなイラストのSVG（都市・製鉄所・基板・港湾・コイル・水素・GD風景・炉の模式図・金属の球） |
| `render_assets.js` | `scenes.js` を `assets/*.png` に書き出す |
| `apply_theme.js` | 書き出した後、テーマの配色をPPTXに書き込む |

- フォントはテーマで `Noto Sans CJK JP` を指定しています。使うPCにこのフォントが無い場合は、無償の Noto Sans CJK JP（または Noto Sans JP）をインストールしてください。PowerPointの［デザイン］→［バリエーション］→［フォント］からテーマフォントを Yu Gothic / メイリオ に一括で切り替えることもできます。
- 章の区切りは `content.js` の `chapterOf()` で変えられます。
- 【要確認】の表示は `build.js` の各スライドにある `foot(s, "...", "要確認の内容")` の第3引数です。確認が済んだら、その引数を削除して再生成してください。
