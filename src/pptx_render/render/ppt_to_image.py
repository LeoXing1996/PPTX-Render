import base64
import os
import os.path as osp
from pathlib import Path
from tempfile import TemporaryDirectory

import pythoncom
import win32com.client


def export_pptx_file(pptx_path: str) -> dict:
    with TemporaryDirectory() as tmp_dir:
        images_output_dir = osp.join(tmp_dir)
        os.makedirs(images_output_dir, exist_ok=True)

        pythoncom.CoInitialize()
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = 1
        try:
            ppt = powerpoint.Presentations.Open(
                osp.abspath(pptx_path), WithWindow=False
            )
            ppt.Export(images_output_dir, "PNG")
            ppt.Close()
        except Exception as exp:
            return {
                "output": "",
                "error": f"Error converting PPTX to images: {str(exp)}",
            }

        finally:
            powerpoint.Quit()
            pythoncom.CoUninitialize()

        try:
            image_file = sorted(Path(images_output_dir).glob("*.PNG"))[0]
            with open(image_file, "rb") as img_file:
                img_bytes = img_file.read()
                img_base64 = base64.b64encode(img_bytes).decode("utf-8")
            return {"output": img_base64, "error": ""}
        except Exception as exp:
            return {"output": "", "error": f"Error reading image file: {str(exp)}"}


if __name__ == "__main__":
    from argparse import ArgumentParser
    import base64
    from io import BytesIO
    from PIL import Image

    parser = ArgumentParser()
    parser.add_argument("--pptx")
    parser.add_argument("--save-path", type=str)
    args = parser.parse_args()

    # print(convert_ppt_to_image_v2(args.pptx))
    result = export_pptx_file(args.pptx)
    img_base64 = result["output"]
    print(result["output"])
    if args.save_path is not None:
        image_stream = base64.b64decode(img_base64)
        img = Image.open(BytesIO(image_stream))
        img.save(args.save_path)
        print(f"Save to {args.save_path}")
