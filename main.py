import os
import io
from typing import List
from fastapi import FastAPI, HTTPException, Depends, status, File, UploadFile, Form
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from pydantic import BaseModel
from google import genai
from google.genai import types
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
            headers={"WWW-Authenticate": "Basic"}, # 💡 フロントで制御するため、ここでのポップアップ強制はフロントのJSで回避します
        )
    return credentials.username

try:
    client = genai.Client()
except Exception as e:
    print(f"Geminiの初期化に失敗しました: {e}")

class DownloadRequest(BaseModel):
    title: str
    result: str

class EditRequest(BaseModel):
    title: str
    current_result: str
    edit_instruction: str

# 💡 修正：誰でもインデックスページ自体は開けるようにする（ポップアップを防止）
@app.get("/", response_class=HTMLResponse)
async def read_index():
    with open("index.html", "r", encoding="utf-8") as f:
        return f.read()

# 🔑 新機能：ログイン画面で入力されたID/パスが正しいかチェックするAPI
@app.get("/auth-check")
async def auth_check(username: str = Depends(authenticate)):
    return {"status": "success", "username": username}

# 📝 新規レポート作成API
@app.post("/process-report")
async def process_report(
    title: str = Form(...),
    report_type: str = Form(...),
    instruction: str = Form(""),
    files: List[UploadFile] = File(...),
    username: str = Depends(authenticate)
):
    ALLOWED_MIMETYPES = ["application/pdf", "image/jpeg", "image/png", "image/webp"]
    if len(files) == 1 and files[0].filename == "":
        raise HTTPException(status_code=400, detail="参考資料ファイルを1つ以上選択してください。")

    for file in files:
        if file.content_type not in ALLOWED_MIMETYPES:
            raise HTTPException(status_code=400, detail=f"ファイル「{file.filename}」は未対応の形式です。")
    
    if report_type == "experiment":
        type_prompt = """
        【レポート構成ルール（電気電子工学実験用）】
        学術的かつ技術的に厳密な表現（〜である調）を使い、以下の構成で詳細に執筆してください。
        ### 1. 実験目的
        ### 2. 実験原理（理論式、回路の基本動作、数式や法則の背景などを含めてください）
        ### 3. 実験方法および回路構成（測定手順や使用機器の注意点を肉付けしてください）
        ### 4. 実験結果（添付資料のデータやグラフの数値を反映・整理してください）
        ### 5. 考察（結果から得られた知見、理論値との誤差の原因、電気工学的な視点での深い分析を【極めて詳細に】書いてください）
        ### 6. 結論
        """
    else:
        type_prompt = """
        【レポート構成ルール（電気応用用）】
        最新の技術動向や社会実装の視点を含め、以下の構成で分かりやすく、かつ専門的に執筆してください。
        ### 1. 総合概要 (Executive Summary)
        ### 2. 技術的背景および基本原理
        ### 3. 現在の主な応用例・社会実装動向（Google検索の最新データをここに強く反映してください）
        ### 4. 現状の技術的課題とその対策
        ### 5. 今後の展望および提言
        """

    gemini_files = []
    try:
        for file in files:
            file_bytes = await file.read()
            gemini_file = client.files.upload(
                file=io.BytesIO(file_bytes),
                config=types.UploadFileConfig(mime_type=file.content_type)
            )
            gemini_files.append(gemini_file)
        
        prompt = f"""
        あなたはプロフェッショナルな電気工学アナリスト兼リサーチライターです。
        提供されたすべての添付ファイルを重要な参考資料として分析し、指定されたタイトルのレポートを執筆してください。
        必要に応じて【Google検索機能（クローリング機能）】を自ら発動し、データの裏付けや最新の技術情報を補完してください。
        
        【作成するレポートのタイトル】: {title}
        【追加の作成指示・メモ】: {instruction}
        {type_prompt}
        
        【超重要ルール】
        ・出力は自動的にWordファイルに変換されるため、AIのメタ発言、前置き、結びの挨拶は【絶対に】含めず、純粋なレポートの章と本文だけを出力してください。
        """
        
        contents = gemini_files + [prompt]
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=contents,
            config=types.GenerateContentConfig(tools=[{"google_search": {}}])
        )
        
        for g_file in gemini_files:
            try: client.files.delete(name=g_file.name)
            except: pass
        
        return {"status": "success", "result": response.text}
        
    except Exception as e:
        for g_file in gemini_files:
            try: client.files.delete(name=g_file.name)
            except: pass
        raise HTTPException(status_code=500, detail=str(e))

# 🛠️ 追加修正API
@app.post("/edit-report")
async def edit_report(request: EditRequest, username: str = Depends(authenticate)):
    try:
        prompt = f"""
        あなたは提出前のレポートを極限までブラッシュアップする優秀な校正・編集者です。
        現在作成されている以下の【現在のレポート】に対して、ユーザーから【修正・追加指示】が届きました。
        
        現在の構成や本文をベースにしつつ、指示された内容を完全に反映した「修正後の最新レポート」を再出力してください。
        必要であれば【Google検索機能（クローリング機能）】を使って追加の専門知識や数式、データを調査して肉付けしてください。
        
        【レポートタイトル】: {request.title}
        【現在のレポート】:
        {request.current_result}
        【ユーザーからの修正・追加指示】:
        {request.edit_instruction}
        
        【超重要ルール】
        ・出力はそのままWordに変換されます。案内文や「修正しました」といった前置きは【絶対に】入れないでください。
        """
        
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=[prompt],
            config=types.GenerateContentConfig(tools=[{"google_search": {}}])
        )
        
        return {"status": "success", "result": response.text}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# 📝 Wordファイルの生成・ダウンロード
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