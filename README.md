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
    import requests
    from PIL import Image
    from io import BytesIO

    def pptx_to_image(url, file_path):
        with open(file_path, "rb") as f:
            files = {"files": (file_path, f)}
            response = requests.post(url, files=files)

        if response.status_code != 200:
            print(response.json())
            raise Exception(f"Request failed with status code {response.status_code}")

        image_stream = BytesIO(response.content)
        img = Image.open(image_stream)
        img.save('test_rendered.png')  # Save the image for verification


    pptx_to_image(
        url='http://localhost:14514/render',
        file_path='autosid/output/cv-v1-3b/0_label.pptx'
    )
   ```

    Then you should see the rendered image saved as `test_rendered.png`.
