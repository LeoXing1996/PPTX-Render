# PPTX-Render

## Server

```bash
python .\src\pptx_render\main.py  # use "\" for windows
```

## Client

1. Start a reverse proxy on your local machine, for example:

   ```bash
   ssh -p 30176 -R 0.0.0.0:14514:localhost:14514 xingzhening@47.100.57.163 -i C:\Users\xingzhening\.ssh\ali
   ```

    * `-p 30176` is the port for remote server (e.g., your DSW)
    * `-R` denotes to start a reverse proxy
    * `0.0.0.0:14514` is the address and port for the remote server
    * `localhost:14514` is the address and port for your local machine
    * `xingzhening@47.100.57.163` is the username and IP address of the remote server
    * `-i` is the path to your private key file, can be omitted if you are using password authentication

2. Request the PPTX file from the remote server:

   ```python
    import base64
    import os.path as osp
    import time
    from io import BytesIO

    import requests
    from PIL import Image


    def pptx_to_image(url, file_list, max_batch_size: int = 8):
        files = []
        for file_path in file_list:
            filename = osp.basename(file_path)
            with open(file_path, "rb") as f:
                files.append(
                    (
                        "files",
                        (filename, f.read(), "application/vnd.openxmlformats-officedocument.presentationml.presentation"),
                    )
                )

        for s_idx in range(0, len(files), max_batch_size):
            e_idx = min(s_idx + max_batch_size, len(files))
            batch_files = files[s_idx:e_idx]
            response = requests.post(url, files=batch_files)
            results = response.json()
            for result in results:
                idx = result["idx"]
                img_bytes = result["bytes"]
                error = result["error"]

                if img_bytes is not None:
                    image_stream = base64.b64decode(img_bytes)
                    img = Image.open(BytesIO(image_stream))
                    img.save(f"test_rendered_{s_idx + idx}.png")  # Save the image for verification
                else:
                    print(f"Error processing file {idx}: {error}")

            print(f'Processed files from index {s_idx} to {e_idx - 1}')


    file_template = "autosid/output/cv-v1-7b-small/{idx}_label.pptx"
    index_list = [idx for idx in range(32) if idx not in [14, 25]]
    file_list = [file_template.format(idx=idx) for idx in index_list]
    s_time = time.time()
    pptx_to_image(
        url="http://localhost:14514/render",
        file_list=file_list,
    )

    print("Total time taken:", time.time() - s_time)
   ```

    Then you should see the rendered image saved as `test_rendered.png`.
