import os
from io import BytesIO
import openai                  # for handling error types
from datetime import datetime  # for formatting date returned with images
import base64                  # for decoding images if recieved in the reply
import requests                # for downloading images from URLs
from PIL import Image          # pillow, for processing image types
import tkinter as tk           # for GUI thumbnails of what we got
from PIL import ImageTk        # for GUI thumbnails of what we got

def old_package(version, minimum):  # Block old openai python libraries before today's
    version_parts = list(map(int, version.split(".")))
    minimum_parts = list(map(int, minimum.split(".")))
    return version_parts < minimum_parts

if old_package(openai.__version__, "1.2.3"):
    raise ValueError(f"Error: OpenAI version {openai.__version__}"
                     " is less than the minimum version 1.2.3\n\n"
                     ">>You should run 'pip install --upgrade openai')")

from openai import OpenAI
from dotenv import load_dotenv
load_dotenv()
api_key = os.environ.get("OPENAI_API_KEY")
client = OpenAI(api_key=api_key) # will use environment variable "OPENAI_API_KEY"

prompt = (
 "Subject: ballet dancers posing on a beam. "  # use the space at end
 "Style: romantic impressionist painting."     # this is implicit line continuation
)

image_params = {
 "model": "dall-e-2",  # Defaults to dall-e-2
 "n": 1,               # Between 2 and 10 is only for DALL-E 2
 "size": "1024x1024",  # 256x256, 512x512 only for DALL-E 2 - not much cheaper
 "prompt": prompt,     # DALL-E 3: max 4000 characters, DALL-E 2: max 1000
 "user": "myName",     # pass a customer ID to OpenAI for abuse monitoring
}

## -- You can uncomment the lines below to include these non-default parameters --

image_params.update({"response_format": "b64_json"})  # defaults to "url" for separate download

## -- DALL-E 3 exclusive parameters --
#image_params.update({"model": "dall-e-3"})  # Upgrade the model name to dall-e-3
#image_params.update({"size": "1792x1024"})  # 1792x1024 or 1024x1792 available for DALL-E 3
#image_params.update({"quality": "hd"})      # quality at 2x the price, defaults to "standard" 
#image_params.update({"style": "natural"})   # defaults to "vivid"

# ---- START
# here's the actual request to API and lots of error catching
try:
    images_response = client.images.generate(**image_params)
except openai.APIConnectionError as e:
    print("Server connection error: {e.__cause__}")  # from httpx.
    raise
except openai.RateLimitError as e:
    print(f"OpenAI RATE LIMIT error {e.status_code}: (e.response)")
    raise
except openai.APIStatusError as e:
    print(f"OpenAI STATUS error {e.status_code}: (e.response)")
    raise
except openai.BadRequestError as e:
    print(f"OpenAI BAD REQUEST error {e.status_code}: (e.response)")
    raise
except Exception as e:
    print(f"An unexpected error occurred: {e}")
    raise

# make a file name prefix from date-time of response
images_dt = datetime.utcfromtimestamp(images_response.created)
img_filename = images_dt.strftime('DALLE-%Y%m%d_%H%M%S')  # like 'DALLE-20231111_144356'

# get the prompt used if rewritten by dall-e-3, null if unchanged by AI
revised_prompt = images_response.data[0].revised_prompt

# get out all the images in API return, whether url or base64
# note the use of pydantic "model.data" style reference and its model_dump() method
image_url_list = []
image_data_list = []
for image in images_response.data:
    image_url_list.append(image.model_dump()["url"])
    image_data_list.append(image.model_dump()["b64_json"])

# Initialize an empty list to store the Image objects
image_objects = []

# Check whether lists contain urls that must be downloaded or b64_json images
if image_url_list and all(image_url_list):
    # Download images from the urls
    for i, url in enumerate(image_url_list):
        while True:
            try:
                print(f"getting URL: {url}")
                response = requests.get(url)
                response.raise_for_status()  # Raises stored HTTPError, if one occurred.
            except requests.HTTPError as e:
                print(f"Failed to download image from {url}. Error: {e.response.status_code}")
                retry = input("Retry? (y/n): ")  # ask script user if image url is bad
                if retry.lower() in ["n", "no"]:  # could wait a bit if not ready
                    raise
                else:
                    continue
            break
        image_objects.append(Image.open(BytesIO(response.content)))  # Append the Image object to the list
        image_objects[i].save(f"{img_filename}_{i}.png")
        print(f"{img_filename}_{i}.png was saved")
elif image_data_list and all(image_data_list):  # if there is b64 data
    # Convert "b64_json" data to png file
    for i, data in enumerate(image_data_list):
        image_objects.append(Image.open(BytesIO(base64.b64decode(data))))  # Append the Image object to the list
        image_objects[i].save(f"{img_filename}_{i}.png")
        print(f"{img_filename}_{i}.png was saved")
else:
    print("No image data was obtained. Maybe bad code?")

## -- extra fun: pop up some thumbnails in your GUI if you want to see what was saved

if image_objects:
    # Create a new window for each image
    for i, img in enumerate(image_objects):
        # Resize image if necessary
        if img.width > 512 or img.height > 512:
            img.thumbnail((512, 512))  # Resize while keeping aspect ratio

        # Create a new tkinter window
        window = tk.Tk()
        window.title(f"Image {i}")

        # Convert PIL Image object to PhotoImage object
        tk_image = ImageTk.PhotoImage(img)

        # Create a label and add the image to it
        label = tk.Label(window, image=tk_image)
        label.pack()

        # Run the tkinter main loop - this will block the script until images are closed
        window.mainloop()