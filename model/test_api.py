import json
import requests
import io
import base64
from PIL import Image
 
url = "https://657dbd09.r11.vip.cpolar.cn"
 
prompt = "dog"
negative_prompt = ""
 
payload = {
 
    # 模型设置
    "override_settings":{
          "sd_model_checkpoint": "[萌二次元]131-half.safetensors",
        #   "sd_vae": "animevae.pt",
          "CLIP_stop_at_last_layers": 2,
    },
 
    # 基本参数
    "prompt": prompt,
    "negative_prompt": negative_prompt,
    "steps": 35,
    "sampler_name": "Euler a",
    "width": 512,
    "height": 512,
    "batch_size": 1,
    "n_iter": 1,
    "seed": 1,
    "CLIP_stop_at_last_layers": 2,
 
    # 面部修复 face fix
    "restore_faces": False,
 
    #高清修复 highres fix
    # "enable_hr": True,
    # "denoising_strength": 0.4,
    # "hr_scale": 2,
    # "hr_upscaler": "Latent",
 
}
 
response = requests.post(url=f'{url}/sdapi/v1/txt2img', json=payload)
r = response.json()
image = Image.open(io.BytesIO(base64.b64decode(r['images'][0])))
 
image.show()
image.save('output.png')