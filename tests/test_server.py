from fastapi import testclient
import pptx_render
import pptx_render.main
from PIL import Image
from io import BytesIO

client = testclient.TestClient(pptx_render.main.app)


def test_render():
    files = {
        "pptx_file": "tests/data/test.pptx",
    }
    file_path = 'tests/data/test.pptx'
    with open(file_path, 'rb') as f:
        files = {"files": (file_path, f)}
        response = client.post('/render', files=files)
    if response.status_code != 200:
        print(response.json())
        raise Exception(f"Request failed with status code {response.status_code}")
    image_stream = BytesIO(response.content)
    img = Image.open(image_stream)
    img.save('tests/data/test_rendered.png')  # Save the image for verification
