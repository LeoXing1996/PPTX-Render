import io
import os.path as osp
from pptx_render.render.ppt_to_image import convert_ppt_to_images
import requests

def test_render():
    test_data = osp.join(__file__, "..", "data", "test.pptx")
    img_bytes = convert_ppt_to_images(test_data, output_dir="./outputs")
    img = io.BytesIO()


if __name__ == "__main__":
    test_render()
