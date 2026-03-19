import os
from PIL import Image, ImageFont, ImageDraw
import customtkinter as ctk

# Material Design Icons map
MDI = {
    "account_balance": "\ue84f",
    "article": "\uef42",
    "assignment": "\ue85d",
    "autorenew": "\ue863",
    "bolt": "\uea0b",
    "business_center": "\ueb3f",
    "check_circle": "\ue86c",
    "close": "\ue5cd",
    "delete": "\ue872",
    "description": "\ue873",
    "folder": "\ue2c7",
    "info": "\ue88e",
    "local_pharmacy": "\ue550",
    "lock": "\ue897",
    "login": "\uea77",
    "medical_services": "\uf109",
    "pause": "\ue034",
    "person": "\ue7fd",
    "play_arrow": "\ue037",
    "save": "\ue161",
    "search": "\ue8b6",
    "settings": "\ue8b8",
    "update": "\ue923",
    "upload_file": "\ue9fc"
}

class IconLib:
    def __init__(self, font_path):
        self.font_path = font_path
        self._cache = {}

    def get_icon(self, name, size=20, light_color="black", dark_color="white"):
        key = (name, size, light_color, dark_color)
        if key in self._cache:
            return self._cache[key]
        
        unicode_char = MDI.get(name, "\ue838") # default star code
        
        try:
            font = ImageFont.truetype(self.font_path, size)
        except Exception as e:
            print("IconLib Error:", e)
            return None

        # Create Light image
        l_img = Image.new("RGBA", (size, size), (255, 255, 255, 0))
        ImageDraw.Draw(l_img).text((0, 0), unicode_char, font=font, fill=light_color)
        
        # Create Dark image
        d_img = Image.new("RGBA", (size, size), (255, 255, 255, 0))
        ImageDraw.Draw(d_img).text((0, 0), unicode_char, font=font, fill=dark_color)
        
        ctk_icon = ctk.CTkImage(light_image=l_img, dark_image=d_img, size=(size, size))
        self._cache[key] = ctk_icon
        return ctk_icon
