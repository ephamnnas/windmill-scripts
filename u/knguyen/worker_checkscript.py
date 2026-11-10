# requirements:
# pillow==10.4.0

import platform
import os

def main():
    pillow_dir = "/tmp/windmill/cache/python_3_12/pillow==10.4.0/PIL"
    return {
        "platform": platform.platform(),
        "machine": platform.machine(),
        "python_version": platform.python_version(),
        "pil_dir_contents": os.listdir(pillow_dir) if os.path.exists(pillow_dir) else "not found",
    }