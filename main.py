# main.py
# All-in-one: YOLOv3 + MediaPipe Hands gestures (1..5) -> TTS suara perempuan (natural)
# Requirements: Python 3.10 (for mediapipe), packages: mediapipe, opencv-python, torch, pywin32 or pyttsx3, numpy, pickle

import os
import time
import threading
import random

import cv2
import mediapipe as mp
import torch
from torch.autograd import Variable
import numpy as np
import pickle as pkl

# YOLO related imports (your repo must have darknet.py, util.py, preprocess.py)
from darknet import Darknet
from util import write_results, load_classes
from preprocess import letterbox_image

# -------------------- TTS: Prefer Windows female SAPI, fallback pyttsx3 --------------------
tts_engine = None
tts_mode = None  # "sapi" or "pyttsx3" or "print"

def init_tts():
    global tts_engine, tts_mode
    # Try SAPI (win32com)
    try:
        import win32com.client as wincl
        sapi = wincl.Dispatch("SAPI.SpVoice")
        # Choose female voice if available
        try:
            voices = sapi.GetVoices()
            female_voice = None
            for i in range(voices.Count):
                name = voices.Item(i).GetDescription().lower()
                if "zira" in name or "female" in name or "hazel" in name or "helena" in name or "helen" in name:
                    female_voice = voices.Item(i)
                    break
            if female_voice:
                sapi.Voice = female_voice
            else:
                # try any voice whose description contains "female"
                for i in range(voices.Count):
                    if "female" in voices.Item(i).GetDescription().lower():
                        sapi.Voice = voices.Item(i)
                        break
        except Exception:
            pass
        sapi.Rate = -1
        try:
            sapi.Volume = 100
        except Exception:
            pass
        tts_engine = sapi
        tts_mode = "sapi"
        print("TTS: using Windows SAPI (female if available).")
        return
    except Exception:
        pass

    # Fallback to pyttsx3 (cross-platform)
    try:
        import pyttsx3
        engine = pyttsx3.init()
        # try to pick female voice
        try:
            voices = engine.getProperty('voices')
            for v in voices:
                name = v.name.lower()
                if "female" in name or "zira" in name or "hazel" in name or "frau" in name:
                    engine.setProperty('voice', v.id)
                    break
        except Exception:
            pass
        engine.setProperty('rate', 150)
        engine.setProperty('volume', 1.0)
        tts_engine = engine
        tts_mode = "pyttsx3"
        print("TTS: using pyttsx3 fallback.")
        return
    except Exception:
        pass

    # Last fallback: print only
    tts_engine = None
    tts_mode = "print"
    print("TTS: no engine available, speech will be printed.")

def tts_speak(text):
    """Non-blocking TTS. Uses chosen engine or prints."""
    if tts_mode == "sapi":
        # use thread to avoid blocking
        threading.Thread(target=lambda: tts_engine.Speak(text), daemon=True).start()
    elif tts_mode == "pyttsx3":
        def job():
            tts_engine.say(text)
            tts_engine.runAndWait()
        threading.Thread(target=job, daemon=True).start()
    else:
        # print fallback
        print("[TTS]", text)

# initialize TTS on import
init_tts()

# -------------------- MediaPipe Hands setup --------------------
mp_hands = mp.solutions.hands
mp_draw = mp.solutions.drawing_utils
hands_detector = mp_hands.Hands(max_num_hands=1, min_detection_confidence=0.6, min_tracking_confidence=0.5)

# -------------------- YOLO setup --------------------
CFG_FILE = "cfg/yolov3.cfg"
WEIGHTS_FILE = "yolov3.weights"
COCO_NAMES = "data/coco.names"
PALETTE_FILE = "pallete"  # optional

if not os.path.exists(CFG_FILE) or not os.path.exists(WEIGHTS_FILE) or not os.path.exists(COCO_NAMES):
    print("WARNING: YOLO files missing (cfg/weights/names). YOLO will error if run.")

classes = load_classes(COCO_NAMES)
try:
    colors = pkl.load(open(PALETTE_FILE, "rb"))
except Exception:
    colors = None

model = Darknet(CFG_FILE)
model.load_weights(WEIGHTS_FILE)
model.net_info["height"] = 160
INP_DIM = int(model.net_info["height"])
model.eval()
CUDA = torch.cuda.is_available()
if CUDA:
    model.cuda()
    torch.backends.cudnn.enabled = True
    torch.backends.cudnn.benchmark = True

# -------------------- helper: prep image --------------------
def prep_image(img, inp_dim):
    orig_im = img
    dim = orig_im.shape[1], orig_im.shape[0]
    img = letterbox_image(orig_im, (inp_dim, inp_dim))
    img_ = img[:, :, ::-1].transpose((2, 0, 1)).copy()
    img_ = torch.from_numpy(img_).float().div(255.0).unsqueeze(0)
    return img_, orig_im, dim

# -------------------- finger counting --------------------
def count_fingers_from_landmarks(hand_landmarks):
    """Return number of fingers up (0..5). Heuristic using tip vs pip positions."""
    lm = hand_landmarks.landmark
    fingers = []

    # Thumb: check x-direction vs previous joint (simple heuristic)
    try:
        if lm[4].x < lm[3].x:
            fingers.append(1)
        else:
            fingers.append(0)
    except Exception:
        fingers.append(0)

    # other 4 fingers (index..pinky): tip y < pip y => up
    tips = [8, 12, 16, 20]
    for t in tips:
        try:
            if lm[t].y < lm[t - 2].y:
                fingers.append(1)
            else:
                fingers.append(0)
        except Exception:
            fingers.append(0)
    return sum(fingers)

# -------------------- Gesture -> phrase mapping --------------------
gesture_phrases = {
    1: "Halo selamat pagi",
    2: "Perkenalkan saya Febri Silaen",
    3: "Senang bertemu dengan anda",
    4: "Salam perubahan",
    5: "Sampai jumpa"
}

# cooldowns
last_spoken_gesture = None
last_spoken_time = 0.0
gesture_cooldown = 2.0  # seconds

# -------------------- main realtime loop --------------------
def main(camera_id=0):
    global last_spoken_gesture, last_spoken_time

    cap = cv2.VideoCapture(camera_id)
    if not cap.isOpened():
        print("ERROR: cannot open camera", camera_id)
        return

    print("Starting. Press 'q' to quit.")
    while True:
        ret, frame = cap.read()
        if not ret:
            break

        fh, fw = frame.shape[:2]
        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)

        # --- Mediapipe hands ---
        finger_count = -1
        results = hands_detector.process(rgb)
        if results.multi_hand_landmarks:
            # use first hand
            for hand_landmarks in results.multi_hand_landmarks:
                mp_draw.draw_landmarks(frame, hand_landmarks, mp_hands.HAND_CONNECTIONS)
                finger_count = count_fingers_from_landmarks(hand_landmarks)
                cv2.putText(frame, f"Jari: {finger_count}", (10, 40),
                            cv2.FONT_HERSHEY_SIMPLEX, 1.0, (0, 200, 0), 2, cv2.LINE_AA)
                break

        # --- Gesture TTS logic ---
        if finger_count != -1:
            now = time.time()
            # Speak when gesture changed or after cooldown
            if (finger_count != last_spoken_gesture) or (now - last_spoken_time > gesture_cooldown):
                phrase = gesture_phrases.get(finger_count)
                if phrase:
                    tts_speak(phrase)
                    last_spoken_gesture = finger_count
                    last_spoken_time = now

        # --- YOLO detection (single-line green boxes) ---
        try:
            img_t, orig_im, dim = prep_image(frame, INP_DIM)
            im_dim = torch.FloatTensor(dim).repeat(1, 2)
            if CUDA:
                img_t = img_t.cuda()
                im_dim = im_dim.cuda()

            output = model(Variable(img_t), CUDA)
            output = write_results(output, 0.6, len(classes), nms=True, nms_conf=0.8)

            if isinstance(output, torch.Tensor) and output.numel() > 0:
                output[:, 1:5] = torch.clamp(output[:, 1:5], 0.0, float(INP_DIM)) / INP_DIM
                output[:, [1, 3]] *= frame.shape[1]
                output[:, [2, 4]] *= frame.shape[0]

                for det in output:
                    cls = int(det[-1])
                    label = classes[cls] if cls < len(classes) else str(cls)
                    x1, y1, x2, y2 = map(int, det[1:5])
                    # draw single-line green rectangle
                    cv2.rectangle(frame, (x1, y1), (x2, y2), (0, 200, 0), 2)
                    cv2.putText(frame, label, (x1, y1-8), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0,200,0), 2)
        except Exception:
            # avoid spamming error prints
            pass

        cv2.imshow("YOLO + Gesture (female TTS)", frame)
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()

if __name__ == "__main__":
    main()
