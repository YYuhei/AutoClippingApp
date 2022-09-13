# 自動切り抜き動画生成プログラム

## 説明
Youtubeをまとめているランキングサイトからスクレイピングをし、  
再生数トップ10の各動画を1分だけ切り抜いて繋ぎ合わせ、動画を作成するプログラム。  
main.pyを実行させると以下のような動画が作成されます。 

![ezgif com-gif-maker](https://user-images.githubusercontent.com/55798139/189537460-a68ec267-3c62-4721-846c-ca83fb2b17a4.gif)

スクレイピングにBeautifulSoup、動画ダウンロードにyoutube-dl(yt-dlp)、  
動画編集にffmpegを使用しています。
