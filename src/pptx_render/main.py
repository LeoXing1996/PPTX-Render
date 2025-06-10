import os.path as osp
import shutil
import subprocess
from tempfile import TemporaryDirectory

import uvicorn
from fastapi import FastAPI, File, UploadFile
from fastapi.responses import JSONResponse

app = FastAPI()


@app.post("/render")
def render_batch(files: list[UploadFile] = File(...)):
    processes = []
    results = []

    with TemporaryDirectory() as tmp_dir:
        for idx, file in enumerate(files):
            try:
                pptx_path = osp.join(tmp_dir, f"pptx_{idx}.pptx")
                with open(osp.join(tmp_dir, pptx_path), "wb") as f:
                    shutil.copyfileobj(file.file, f)
                proc = subprocess.Popen(
                    [
                        "python",
                        "src/pptx_render/render/ppt_to_image.py",
                        "--pptx",
                        pptx_path,
                    ],
                    stdin=subprocess.PIPE,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                )
                error = ""
            except Exception as exp:
                proc = None
                error = f"Error creating subprocess: {exp}"
            finally:
                processes.append((idx, proc, error))
                file.file.close()

        for idx, proc, error in processes:
            if len(error) > 0:
                results.append(
                    {
                        "idx": idx,
                        "bytes": None,
                        "error": error,
                    }
                )
                continue

            stdout, stderr = proc.communicate()
            if proc.returncode != 0:
                results.append(
                    {
                        "idx": idx,
                        "bytes": None,
                        "error": stderr.strip(),
                    }
                )
            else:
                results.append(
                    {
                        "idx": idx,
                        "bytes": stdout,
                        "error": None,
                    }
                )

    return JSONResponse(content=results)


if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=14515, log_level="info")
