# main.py
# YOLO + MediaPipe Hands + Simple Sign-Language (heuristic) -> Bahasa Indonesia + TTS
# Pastikan environment memakai Python 3.10 (mediapipe), dan library sudah terinstal.

import cv2
import mediapipe as mp
import win32com.client as wincl
import time
import torch
from torch.autograd import Variable
from darknet import Darknet
from util import write_results, load_classes
from preprocess import letterbox_image
import random
import pickle as pkl
import numpy as np
import threading
import os

# ------------------- TTS -------------------
speak = wincl.Dispatch("SAPI.SpVoice")
speak.Rate = -1
speak.Volume = 90

def tts_nonblocking(text):
    threading.Thread(target=speak.Speak, args=(text,), daemon=True).start()

# ------------------- Mediapipe Hands -------------------
mp_hands = mp.solutions.hands
hands = mp_hands.Hands(max_num_hands=1,
                       min_detection_confidence=0.6,
                       min_tracking_confidence=0.5)
mp_draw = mp.solutions.drawing_utils

# ------------------- YOLO Setup -------------------
cfgfile = "cfg/yolov3.cfg"
weightsfile = "yolov3.weights"
classes = load_classes("data/coco.names")
# pallete optional; we will draw boxes in green for consistency
try:
    colors = pkl.load(open("pallete", "rb"))
except Exception:
    colors = None

model = Darknet(cfgfile)
model.load_weights(weightsfile)
model.net_info["height"] = 160
inp_dim = int(model.net_info["height"])
model.eval()
CUDA = torch.cuda.is_available()
if CUDA:
    model.cuda()
    import torch.backends.cudnn as cudnn
    cudnn.benchmark = True

# ------------------- Utils -------------------
def prep_image(img, inp_dim):
    orig_im = img
    dim = orig_im.shape[1], orig_im.shape[0]
    img = letterbox_image(orig_im, (inp_dim, inp_dim))
    img_ = img[:, :, ::-1].transpose((2, 0, 1)).copy()
    img_ = torch.from_numpy(img_).float().div(255.0).unsqueeze(0)
    return img_, orig_im, dim

# ------------------- Simple sign-language classifier (heuristic) -------------------
# Input: mediapipe hand_landmarks
# Output: (label_str or None, confidence_estimate)
def classify_sign(hand_landmarks):
    """
    Heuristic rules:
    - open palm (all 5 fingers up) -> "Halo"
    - two fingers (index+middle up) -> "Terima kasih"
    - one finger (only index up) -> "Saya"
    - fist (0 fingers) -> "Diam"
    - thumb up (thumb extended, others folded) -> "Bagus"
    Returns label string or None.
    """
    try:
        lm = hand_landmarks.landmark
        # finger tip indices: [4-thumb, 8-index, 12-middle, 16-ring, 20-pinky]
        tips = [4, 8, 12, 16, 20]
        fingers_up = [0]*5

        # Determine handness roughly by comparing wrist x vs index_mcp x could be used,
        # but for simplicity use vertical comparisons for 4 fingers.
        # For thumb, use x-direction because thumb extends sideways.
        # Note: mediapipe coords: x left->right, y top->bottom

        # Thumb: compare tip x with ip x (tip < ip for right hand when extended to left)
        try:
            if lm[tips[0]].x < lm[tips[0]-1].x:   # heuristic for "extended"
                fingers_up[0] = 1
            else:
                fingers_up[0] = 0
        except Exception:
            fingers_up[0] = 0

        # Other fingers: tip y < pip y => finger up
        for i, tid in enumerate(tips[1:], start=1):
            try:
                if lm[tid].y < lm[tid-2].y:
                    fingers_up[i] = 1
                else:
                    fingers_up[i] = 0
            except Exception:
                fingers_up[i] = 0

        total = sum(fingers_up)

        # Heuristics for thumb-up: thumb up and others down
        if fingers_up[0] == 1 and sum(fingers_up[1:]) == 0:
            return "Bagus", 0.85

        # Open palm
        if total == 5:
            return "Halo", 0.9

        # Two-finger V (index+middle)
        if fingers_up[1] == 1 and fingers_up[2] == 1 and fingers_up[3] == 0 and fingers_up[4] == 0:
            return "Terima kasih", 0.85

        # One finger (only index)
        if fingers_up[1] == 1 and fingers_up[2] == 0 and fingers_up[3] == 0 and fingers_up[4] == 0 and fingers_up[0] == 0:
            return "Saya", 0.8

        # Fist
        if total == 0:
            return "Diam", 0.8

        # Fallback none
        return None, 0.0

    except Exception:
        return None, 0.0

# ------------------- Gesture detection cooldown & display -------------------
last_sign = None
last_sign_time = 0.0
sign_cooldown = 2.0  # seconds between announcing same sign

# ------------------- Main loop -------------------
def main():
    global last_sign, last_sign_time

    cap = cv2.VideoCapture(0)
    if not cap.isOpened():
        print("Error: kamera tidak terbuka")
        return

    print("Menjalankan: tekan 'q' untuk keluar")

    while True:
        ret, frame = cap.read()
        if not ret:
            break

        h, w, _ = frame.shape
        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)

        # --- Mediapipe hands detection ---
        results = hands.process(rgb)
        detected_sign = None
        detected_conf = 0.0

        if results.multi_hand_landmarks:
            for hand_landmarks in results.multi_hand_landmarks:
                mp_draw.draw_landmarks(frame, hand_landmarks, mp_hands.HAND_CONNECTIONS)

                sign_label, conf = classify_sign(hand_landmarks)
                if sign_label is not None:
                    detected_sign = sign_label
                    detected_conf = conf

                # show landmarks small debug (optional)
                # for i, lm in enumerate(hand_landmarks.landmark):
                #     cx, cy = int(lm.x * w), int(lm.y * h)
                #     cv2.circle(frame, (cx, cy), 1, (255, 0, 0), -1)

        # announce sign with cooldown and overlay
        if detected_sign is not None:
            now = time.time()
            # Avoid repeating same sign too frequently
            if detected_sign != last_sign or (now - last_sign_time) > sign_cooldown:
                # TTS announce (Bahasa Indonesia)
                msg = ""
                # Map label to natural phrase (in Indonesian)
                if detected_sign == "Halo":
                    msg = "Halo"
                elif detected_sign == "Terima kasih":
                    msg = "Terima kasih"
                elif detected_sign == "Saya":
                    msg = "Saya"
                elif detected_sign == "Diam":
                    msg = "Diam"
                elif detected_sign == "Bagus":
                    msg = "Bagus"
                else:
                    msg = detected_sign

                tts_nonblocking(msg)
                last_sign = detected_sign
                last_sign_time = now

            # overlay text (top-left)
            cv2.putText(frame, f"Isyarat: {detected_sign} ({int(detected_conf*100)}%)", (10,30),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.9, (0,200,0), 2, cv2.LINE_AA)

        # --- YOLO detection (kept minimal) ---
        try:
            img, orig_im, dim = prep_image(frame, inp_dim)
            im_dim = torch.FloatTensor(dim).repeat(1, 2)
            if CUDA:
                img = img.cuda()
                im_dim = im_dim.cuda()

            output = model(Variable(img), CUDA)
            output = write_results(output, 0.6, 80, nms=True, nms_conf=0.8)

            if isinstance(output, torch.Tensor) and output.numel() > 0:
                # scale bboxes
                output[:, 1:5] = torch.clamp(output[:, 1:5], 0.0, float(inp_dim)) / inp_dim
                output[:, [1, 3]] *= frame.shape[1]
                output[:, [2, 4]] *= frame.shape[0]

                for detection in output:
                    cls = int(detection[-1])
                    label = classes[cls]
                    # draw single-line green box
                    x1, y1, x2, y2 = map(int, detection[1:5])
                    cv2.rectangle(frame, (x1, y1), (x2, y2), (0,255,0), 2)
                    cv2.putText(frame, label, (x1, y1-6), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0,255,0), 2)

        except Exception as e:
            # you may print error once during debugging; avoid spamming
            # print("YOLO ERROR:", e)
            pass

        cv2.imshow("SignLang + YOLO", frame)
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()

if __name__ == "__main__":
    main()
