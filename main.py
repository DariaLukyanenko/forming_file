from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import StreamingResponse
import json
import xlsxwriter
import io
from parse_ogrn_nalog import scrape_ogrn_info
from pydantic import BaseModel


app = FastAPI()


class OGRNRequest(BaseModel):
    ogrn: str


@app.post("/upload_file_forming")
async def upload_file(file: UploadFile = File(...)):
    if not file.filename.endswith('.json'):
        raise HTTPException(status_code=400, detail="Unsupported file format")

    contents = await file.read()
    users_data = json.loads(contents)

    # Create an Excel file in memory
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet()

    # Write headers in the first row
    headers = users_data[0].keys()
    for col, header in enumerate(headers):
        worksheet.write(0, col, header)

    # Write each user's data in the subsequent rows
    for row, user in enumerate(users_data, start=1):
        for col, (key, value) in enumerate(user.items()):
            worksheet.write(row, col, value)

    # Set column widths
    for col, header in enumerate(headers):
        max_width = max(len(str(header)), max(len(str(user[header])) for user in users_data))
        worksheet.set_column(col, col, max_width + 2)  # +2 for padding

    worksheet.autofilter(0, 0, len(users_data), len(headers) - 1)

    workbook.close()
    output.seek(0)

    return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             headers={"Content-Disposition": "attachment;filename=all_users_data.xlsx"})


@app.post("/get-info_ogrn")
def get_info(request: OGRNRequest):
    try:
        data = scrape_ogrn_info(request.ogrn)
        if data:
            return data
        else:
            raise HTTPException(status_code=500, detail="Failed to retrieve data")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8000)
