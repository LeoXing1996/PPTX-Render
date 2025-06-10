import base64
from io import BytesIO

from fastapi import testclient
from PIL import Image

import pptx_render
import pptx_render.main

client = testclient.TestClient(pptx_render.main.app)


def test_render_batch():
    file_path_list = [
        "tests/data/test.pptx",
        "tests/data/test.pptx",
        "tests/data/test.pptx",
        "tests/data/test.pptx",
        "tests/data/test.pptx",
        "tests/data/test.pptx",
    ]
    files = [
        ("files", (file_path, open(file_path, "rb"))) for file_path in file_path_list
    ]
    response = client.post("/render", files=files)

    if response.status_code != 200:
        print(response.json())
        raise Exception(f"Request failed with status code {response.status_code}")

    results = response.json()
    for result in results:
        idx = result["idx"]
        img_bytes = result["bytes"]
        error = result["error"]

        if img_bytes is not None:
            image_stream = base64.b64decode(img_bytes)
            img = Image.open(BytesIO(image_stream))
            img.save(f"tests/data/test_rendered_v2_{idx}.png")
        else:
            print(f"Error processing file {idx}: {error}")
