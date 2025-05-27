import os
import os.path as osp
from pathlib import Path
from threading import Lock

import win32com.client
import pythoncom


pptx_lock = Lock()


def convert_ppt_to_images(pptx_path: str, output_dir: str = None) -> str:
    with pptx_lock:
        try:
            pythoncom.CoInitialize()
            if output_dir is None:
                output_dir = os.path.join("outputs", Path(pptx_path).stem)
            output_dir = osp.abspath(output_dir)
            os.makedirs(output_dir, exist_ok=True)

            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            ppt_app.Visible = 1
            presentation = ppt_app.Presentations.Open(pptx_path, WithWindow=False)

            presentation.Export(output_dir, "PNG")
            presentation.Close()
            ppt_app.Quit()

            image_files = sorted(Path(output_dir).glob("*.PNG"))
            return image_files[0]
        except Exception as e:
            # Ensure cleanup on error
            if "presentation" in locals():
                presentation.Close()
            if "ppt_app" in locals():
                ppt_app.Quit()
            raise Exception(f"Failed to convert PPTX to images: {str(e)}")
        finally:
            # Ensure COM is uninitialized (though handled in endpoint)
            # pythoncom.CoUninitialize()
            pythoncom.CoUninitialize()  # Ensure COM is uninitialized
