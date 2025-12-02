import torch, cv2, random, os, time
import torch.nn as nn
from torch.autograd import Variable
import numpy as np
import pickle as pkl
import argparse
import threading, queue
from torch.multiprocessing import Pool, Process, set_start_method
from util import write_results, load_classes
from preprocess import letterbox_image
from darknet import Darknet
from imutils.video import WebcamVideoStream, FPS

# ---------------- TEXT TO SPEECH -----------------
import win32com.client as wincl
speak = wincl.Dispatch("SAPI.SpVoice")

voices = speak.GetVoices()
preferred_keywords = ["indonesia", "bahasa", "indonesian", "siti", "ayu", "nina"]

voice_found = False
for i, voice in enumerate(voices):
    desc = voice.GetDescription().lower()
    if any(k in desc for k in preferred_keywords):
        speak.Voice = voices.Item(i)
        voice_found = True
        print(f"✓ Suara Indonesia digunakan: {voices.Item(i).GetDescription()}")
        break

if not voice_found:
    speak.Voice = voices.Item(0)
    print(f"✓ Suara default digunakan: {voices.Item(0).GetDescription()}")

speak.Rate = -1
speak.Volume = 90

try:
    torch.multiprocessing.set_start_method('spawn', force=True)
except RuntimeError:
    pass

# ------------------------------------------------------

if torch.cuda.is_available():
    torch.backends.cudnn.enabled = True
    torch.backends.cudnn.benchmark = True
    torch.backends.cudnn.deterministic = True
    torch.set_default_tensor_type('torch.cuda.FloatTensor')

def prep_image(img, inp_dim):
    orig_im = img
    dim = orig_im.shape[1], orig_im.shape[0]
    img = (letterbox_image(orig_im, (inp_dim, inp_dim)))
    img_ = img[:, :, ::-1].transpose((2,0,1)).copy()
    img_ = torch.from_numpy(img_).float().div(255.0).unsqueeze(0)
    return img_, orig_im, dim


# ---------------- SMART SPEAKING --------------------

last_speak_time = 0
last_spoken_object = None
last_spoken_distance = None

speak_cooldown = 5
distance_change_threshold = 20

object_translations = {
    "person": "orang",
    "cell phone": "ponsel",
    "book": "buku",
    "bottle": "botol",
    "cup": "gelas",
    "laptop": "laptop",
    "mouse": "mouse",
    "keyboard": "keyboard",
    "chair": "kursi",
    "couch": "sofa",
    "tv": "televisi"
}

def get_natural_phrase(label, distance_cm):
    obj_name = object_translations.get(label.lower(), label)
    phrases = [
        f"Mendeteksi {obj_name} pada jarak {distance_cm} sentimeter.",
        f"Ada {obj_name} sekitar {distance_cm} sentimeter.",
        f"{obj_name} terdeteksi, jarak {distance_cm} sentimeter.",
    ]
    return random.choice(phrases)

def should_announce(label, distance_cm):
    global last_speak_time, last_spoken_object, last_spoken_distance

    now = time.time()

    if last_spoken_object is None:
        return True
    
    if now - last_speak_time < speak_cooldown:
        return False
    
    if label != last_spoken_object:
        return True
    
    if last_spoken_distance is not None:
        if abs(distance_cm - last_spoken_distance) > distance_change_threshold:
            return True

    return False

# ---------------- CLEAN BOUNDING BOX --------------------

def write(bboxes_row, img, classes, colors=None):
    """
    FIXED CLEAN VERSION:
    - 1 garis hijau saja
    - Tidak ada shadow / outline / layer
    - Text clean
    """
    global last_speak_time, last_spoken_object, last_spoken_distance

    try:
        cls = int(bboxes_row[-1])
        x1, y1, x2, y2 = bboxes_row[1:5].int()

        label = classes[cls]

        # Warna hijau solid
        color = (0, 255, 0)

        # Hitung jarak
        w = x2 - x1
        h = y2 - y1
        distance = (2 * 3.14 * 180) / (float(w) + float(h) * 360) * 1000 + 3
        distance_cm = round(distance * 2.54)

        # Smart announcement
        if should_announce(label, distance_cm):
            message = get_natural_phrase(label, distance_cm)
            threading.Thread(target=speak.Speak, args=(message,), daemon=True).start()
            
            last_speak_time = time.time()
            last_spoken_object = label
            last_spoken_distance = distance_cm

            print(f"🔊 {message}")
        else:
            print(f"📷 {label} {distance_cm} cm")

        # DRAW BOUNDING BOX CLEAN
        cv2.rectangle(img, (int(x1), int(y1)), (int(x2), int(y2)), color, 2)

        # Tambah label
        text = f"{label} ({distance_cm} cm)"
        cv2.putText(
            img, text, (int(x1), int(y1)-8),
            cv2.FONT_HERSHEY_SIMPLEX, 0.6, color, 2, cv2.LINE_AA
        )

    except Exception as e:
        print(f"ERROR in write(): {e}")

    return img


# ---------------- YOLO CLASS --------------------

class ObjectDetection:
    def __init__(self, id):
        print("Initializing camera...")
        self.cap = WebcamVideoStream(src=id).start()
        time.sleep(2)

        self.cfgfile = "cfg/yolov3.cfg"
        self.weightsfile = "yolov3.weights"

        self.confidence = 0.6
        self.nms_thesh = 0.8
        self.num_classes = 80

        self.classes = load_classes("data/coco.names")
        self.colors = None   # REMOVE pallete completely

        print("Loading YOLO model...")
        self.model = Darknet(self.cfgfile)
        self.model.load_weights(self.weightsfile)
        self.model.net_info["height"] = 160
        self.inp_dim = 160

        if torch.cuda.is_available():
            self.model.cuda()
            print("✓ GPU Mode")
        else:
            print("✓ CPU Mode")

        self.model.eval()

        self.width = 1280
        self.height = 720

        print("✓ YOLO loaded")
        print("✓ Voice system ready")

    def main(self):
        q = queue.Queue()

        def frame_capture(q):
            frame = self.cap.read()
            if frame is not None:
                frame = cv2.resize(frame, (self.width, self.height))
                q.put(frame)

        cam = threading.Thread(target=frame_capture, args=(q,))
        cam.start()
        cam.join()

        if q.empty():
            print("Camera error")
            return None

        frame = q.get()
        q.task_done()

        img, orig_im, dim = prep_image(frame, self.inp_dim)
        im_dim = torch.FloatTensor(dim).repeat(1,2)

        if torch.cuda.is_available():
            img = img.cuda()
            im_dim = im_dim.cuda()

        output = self.model(Variable(img), torch.cuda.is_available())
        output = write_results(output, self.confidence, self.num_classes, nms=True, nms_conf=self.nms_thesh)

        if isinstance(output, torch.Tensor) and output.numel() > 0:
            output[:, 1:5] = torch.clamp(output[:, 1:5], 0.0, float(self.inp_dim)) / self.inp_dim
            output[:, [1,3]] *= frame.shape[1]
            output[:, [2,4]] *= frame.shape[0]

            for det in output:
                frame = write(det, frame, self.classes, self.colors)

        ret, jpeg = cv2.imencode(".jpg", frame)
        return jpeg.tobytes() if ret else None


# ---------------- MAIN --------------------

if __name__ == "__main__":
    od = ObjectDetection(0)
    result = od.main()
    if result:
        with open("detected_frame.jpg", "wb") as f:
            f.write(result)
        print("Frame saved ✔")
    else:
        print("No result.")
