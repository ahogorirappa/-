import os
import io
from typing import List
from fastapi import FastAPI, HTTPException, Depends, status, File, UploadFile, Form
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from pydantic import BaseModel
from google import genai
from google.genai import types  # Geminiの各種設定を使うためのライブラリ
from dotenv import load_dotenv
from docx import Document

load_dotenv()
app = FastAPI()
security = HTTPBasic()

def authenticate(credentials: HTTPBasicCredentials = Depends(security)):
    correct_username = os.getenv("WEB_USERNAME", "1")
    correct_password = os.getenv("WEB_PASSWORD", "1")
    if credentials.username != correct_username or credentials.password != correct_password:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="ユーザー名またはパスワードが違います",
            headers={"WWW-Authenticate": "Basic"},
        )
    return credentials.username

try:
    client = genai.Client()
except Exception as e:
    print(f"Geminiの初期化に失敗しました: {e}")

class DownloadRequest(BaseModel):
    title: str
    result: str

@app.get("/", response_class=HTMLResponse)
async def read_index(username: str = Depends(authenticate)):
    with open("index.html", "r", encoding="utf-8") as f:
        return f.read()

# 📝 複数ファイルをまとめて読み込み、Google検索と連動してレポートを作るAPI
@app.post("/process-report")
async def process_report(
    title: str = Form(...),
    instruction: str = Form(""),
    files: List[UploadFile] = File(...),
    username: str = Depends(authenticate)
):
    ALLOWED_MIMETYPES = ["application/pdf", "image/jpeg", "image/png", "image/webp"]
    
    for file in files:
        if file.content_type not in ALLOWED_MIMETYPES:
            raise HTTPException(status_code=400, detail=f"ファイル「{file.filename}」は未対応の形式です。")
    
    gemini_files = []
    try:
        # 1. 画面から届いた複数のファイルをすべてGeminiのサーバーへアップロード
        for file in files:
            file_bytes = await file.read()
            print(f"Geminiへ参考資料「{file.filename}」を安全にアップロード中...")
            
            gemini_file = client.files.upload(
                file=io.BytesIO(file_bytes),
                config=types.UploadFileConfig(mime_type=file.content_type)
            )
            gemini_files.append(gemini_file)
        
        # 2. 調査と清書を一度に行うための最強プロンプト
        prompt = f"""
        あなたはプロフェッショナルな最高峰のリサーチアナリスト兼ライターです。
        提供されたすべての添付ファイルを【重要な参考資料群】として厳密に分析し、それらの情報を総合的に組み合わせて、指定されたタイトルの「新規レポート」を思考・執筆してください。
        
        さらに、レポートの完成度を極限まで高めるため、必要に応じて同時に有効化されている【Google検索機能（クローリング機能）】を自ら発動し、資料内のデータの裏付け、最新の市場動向、関連する統計情報などをリアルタイムに調査して補完してください。
        
        単なる資料のまとめにと度まらず、深い洞察と最新トレンドを掛け合わせた、極めて完成度の高いレポートを作成してください。
        
        【作成するレポートのタイトル】: {title}
        【追加の作成指示・メモ】: {instruction}
        
        【レポートの構成ルール】
        必ず以下の構成（見出し）に沿って、詳細かつ具体的に執筆してください。
        ### 1. 総合概要 (Executive Summary)
        ### 2. 本書の中心となる最重要トピック（3〜5点）
        ### 3. 詳細な分析と解説（検索した最新データや市場背景、複数資料のクロス分析をここに肉付けしてください）
        ### 4. 結論および今後の提言・次のアクション
        
        【超重要ルール】
        ・出力されたテキストは、この後システムによって自動的にWord（.docx）ファイルに変換されます。
        ・そのため、「私は直接ファイルを生成できません」「検索した結果〜」といった、AIとしてのメタな発言や言い訳、案内文は【絶対に】出力に含めないでください。
        ・前置きの挨拶や結びの言葉も一切不要です。純粋なレポートの文章（見出しと本文）だけを出力してください。
        """
        
        print("Geminiが資料を解析し、自動Google検索（クローリング）を併用してレポートを生成中...")
        contents = gemini_files + [prompt]
        
        # 3. 🔍 Google検索ツールを合体させてGeminiを呼び出す（これが本当の贅沢品！）
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=contents,
            config=types.GenerateContentConfig(
                tools=[{"google_search": {}}]  # ← ここでGoogle検索（クローリング）機能をドッキング！
            )
        )
        
        # 4. 終わったらGemini側のファイルを綺麗にお掃除
        for g_file in gemini_files:
            try: client.files.delete(name=g_file.name)
            except: pass
        
        return {"status": "success", "result": response.text}
        
    except Exception as e:
        for g_file in gemini_files:
            try: client.files.delete(name=g_file.name)
            except: pass
        raise HTTPException(status_code=500, detail=str(e))

# 📝 Wordファイル（.docx）の生成・ダウンロード
@app.post("/download/docx")
async def download_docx(request: DownloadRequest, username: str = Depends(authenticate)):
    doc = Document()
    doc.add_heading(request.title, level=1)
    
    for line in request.result.split("\n"):
        if line.startswith("### "):
            doc.add_heading(line.replace("### ", ""), level=3)
        elif line.startswith("## "):
            doc.add_heading(line.replace("## ", ""), level=2)
        elif line.startswith("- ") or line.startswith("* "):
            doc.add_paragraph(line[2:], style='List Bullet')
        else:
            if line.strip():
                doc.add_paragraph(line)
                
    file_io = io.BytesIO()
    doc.save(file_io)
    file_io.seek(0)
    return StreamingResponse(
        file_io,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": "attachment; filename=report.docx"}
    )