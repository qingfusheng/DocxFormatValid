import os.path

from fastapi import FastAPI, HTTPException, UploadFile, File, Response
from fastapi.middleware.cors import CORSMiddleware
import io
import datetime
from PaperFormatDetection import PaperFormatDt
import uvicorn

pfd = PaperFormatDt()
app = FastAPI()
origins = [
    "*"
]

# 添加跨域中间件
app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def generate_filename_with_timestamp(extension=".docx"):
    now = datetime.datetime.now()
    timestamp = now.strftime("%Y%m%d%H%M%S%f")
    filename = f"file_{timestamp}{extension}"
    return filename


# def process_docx_to_pdf(docx_data: bytes) -> io.BytesIO:
#     # 在这里添加处理.docx文件并生成.pdf文件的逻辑
#     # 这里仅将.docx数据返回为.pdf数据的示例
#     # 实际应用中需要使用相应的库将.docx转换为.pdf
#
#     # 这里仅将.docx数据包装成一个假的.pdf文件
#     pdf_data = io.BytesIO(docx_data)
#     pdf_data.name = "output.pdf"
#     return pdf_data

@app.get("/")
async def root():
    return "Hello World"


@app.post("/uplad/")
async def paper_detect(file: UploadFile = File(...)) -> Response:
    # 检查文件扩展名
    if not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Only .docx files are allowed.")
    try:
        # 将上传的.docx文件读取为bytes
        docx_data = await file.read()
        print("上传完成")
        filename = generate_filename_with_timestamp()
        with open(os.path.join(os.path.abspath("../code/"), "DocxFilter", filename), "wb") as f:
            f.write(docx_data)
        pfd.run(filename)
        with open(os.path.join(os.path.abspath("../code/"), "OutputDocxFilter", filename), "rb") as f:
            docx_data = f.read()
        # 返回生成的.docx文件数据
        return Response(content=docx_data, media_type="application/docx")

    except Exception as e:
        # 处理可能的错误
        print(e)
        error_code = 500  # 默认错误代码
        error_detail = str(e)
        if isinstance(e, FileNotFoundError):
            error_code = 404
        elif isinstance(e, Exception):
            # 可根据具体情况添加其他错误处理逻辑
            pass
        raise HTTPException(status_code=error_code, detail=error_detail)


if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=5000, reload=True)
