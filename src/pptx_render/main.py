import io
import pythoncom
import os
import os.path as osp
import shutil
from tempfile import TemporaryDirectory

import uvicorn
from fastapi import FastAPI, File, HTTPException, Response, UploadFile
from PIL import Image

from pptx_render.render.ppt_to_image import convert_ppt_to_images

app = FastAPI()


@app.post("/render")
async def do_render(files: UploadFile = File(...)):
    pptx_name = osp.basename(files.filename)
    print(files)
    if not pptx_name.endswith(".pptx"):
        raise HTTPException(status_code=400, detail=f"Only .pptx files are allowed, input is {pptx_name}")

    with TemporaryDirectory() as tmp_dir:
        pptx_path = osp.join(tmp_dir, pptx_name)
        images_output_dir = osp.join(tmp_dir, "images")
        os.makedirs(images_output_dir, exist_ok=True)

        try:
            with open(pptx_path, "wb") as buffer:
                shutil.copyfileobj(files.file, buffer)
        except Exception as exp:
            raise HTTPException(
                status_code=500, detail=f"Error saving file: {str(exp)}"
            )
        finally:
            files.file.close()

        # Convert PPTX to images
        try:
            img = convert_ppt_to_images(pptx_path, output_dir=images_output_dir)

        except Exception as exp:
            raise HTTPException(
                status_code=500, detail=f"Error converting PPTX to images: {str(exp)}"
            )

        pil_img = Image.open(img)
        img_byte_arr = io.BytesIO()
        pil_img.save(img_byte_arr, format="PNG")
        img_byte_arr.seek(0)

        return Response(content=img_byte_arr.getvalue(), media_type="image/png")


if __name__ == '__main__':
    uvicorn.run(app, host='0.0.0.0', port=14514, log_level="info")
