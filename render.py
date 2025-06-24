import base64
import os
import os.path as osp
from argparse import ArgumentParser
from io import BytesIO
from typing import Optional

import requests
from PIL import Image
from tqdm import trange


def get_args():
    parser = ArgumentParser()
    parser.add_argument(
        "--gen-dir", type=str, help="Path to the generated PPTX or code"
    )
    parser.add_argument("--max-samples", type=int, default=100)
    parser.add_argument(
        "--save-dir", type=str, default="Path to save the comparison results"
    )

    parser.add_argument(
        "--batch-size",
        type=int,
        default=8,
        help="The max batch size to request the render server",
    )
    parser.add_argument(
        "--api", type=str, default="render", help="The name of the render API"
    )
    parser.add_argument(
        "--port", type=int, default="14515", help="The port of the render API"
    )

    return parser.parse_args()


def pptx_to_image(
    url: str,
    file_list: list[str],
    save_dir: str,
    max_batch_size: int = 8,
    src_list: Optional[list[str]] = None,
    save_src: bool = False,
    concat: bool = False,
) -> list[dict]:
    if save_src or concat:
        assert src_list is not None, (
            "'src_list' must be passed if `save_src` or `concat` is True"
        )

    for s_idx in trange(0, len(file_list), max_batch_size):
        e_idx = min(s_idx + max_batch_size, len(file_list))

        # NOTE: dirty checking, but work
        batch_files, save_idx_list = [], []
        for idx, file_path in enumerate(file_list[s_idx:e_idx]):
            gen_save_path = osp.join(save_dir, f"{s_idx + idx}_gen.jpg")
            if osp.exists(gen_save_path):
                continue
            filename = osp.basename(file_path)
            with open(file_path, "rb") as f:
                batch_files.append(
                    (
                        "files",
                        (
                            filename,
                            f.read(),
                            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        ),
                    )
                )

            save_idx_list.append(s_idx + idx)

        if not batch_files:
            print(f"Skip {s_idx} to {e_idx} as all files are already generated.")
            continue

        response = requests.post(url, files=batch_files)
        results = response.json()
        for result, save_idx in zip(results, save_idx_list):
            idx = result["idx"]
            img_bytes = result["bytes"]
            error = result["error"]

            if img_bytes is not None:
                image_stream = base64.b64decode(img_bytes)
                try:
                    img = Image.open(BytesIO(image_stream)).convert("RGB")
                except Exception as Exp:
                    img = None
                    error = f"Error decoding image bytes: {Exp}"
            else:
                img = None

            if img is not None:
                gen_save_path = osp.join(save_dir, f"{save_idx}_gen.jpg")
                img.save(gen_save_path)
            else:
                gen_save_path = osp.join(save_dir, f"{save_idx}_gen.txt")
                with open(gen_save_path, "w") as file:
                    file.write(error)



def test_api(url: str) -> bool:
    response = requests.get(url)
    if response.status_code == 200:
        print("Render API is accessable!")
        return True
    else:
        print("Render API is un-accessable!")
        print("Response:")
        print(response.json())
        return False


def main(args):
    api_url = f"http://127.0.0.1:{args.port}/{args.api}"
    api_test_url = f"http://127.0.0.1:{args.port}/docs"
    if not (test_api(api_test_url)):
        exit()

    gen_dir = args.gen_dir
    gen_samples = list(filter(lambda s: s.endswith(".pptx"), os.listdir(gen_dir)))
    gen_samples.sort()

    # src_dir = args.src_dir
    gen_to_render = []
    for sample in gen_samples:
        gen_pptx = osp.join(gen_dir, sample)
        gen_to_render.append(gen_pptx)

        if args.max_samples != -1 and len(gen_to_render) >= args.max_samples:
            break

    save_dir = args.save_dir
    os.makedirs(save_dir, exist_ok=True)
    pptx_to_image(
        api_url,
        gen_to_render,
        save_dir,
        max_batch_size=args.batch_size,
        src_list=[],
        save_src=args.save_src,
    )


if __name__ == "__main__":
    """
    Test case (use "\" for windows):
    python .\render.py --gen-dir .\cv-v1-vis\  --save-dir .\cv-v1-vis-render --max-samples -1
    """
    args = get_args()
    main(args)
