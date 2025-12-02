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

# Text-to-speech setup dengan suara perempuan
import win32com.client as wincl
speak = wincl.Dispatch("SAPI.SpVoice")

# Pilih suara perempuan (biasanya index 1, tapi bisa berbeda di sistem Anda)
voices = speak.GetVoices()
# Coba cari suara perempuan
for i, voice in enumerate(voices):
    if "female" in voice.GetDescription().lower() or "zira" in voice.GetDescription().lower():
        speak.Voice = voices.Item(i)
        print(f"Menggunakan suara: {voice.GetDescription()}")
        break
else:
    # Jika tidak ada keyword "female", pakai index 1 (biasanya perempuan)
    if len(voices) > 1:
        speak.Voice = voices.Item(1)
        print(f"Menggunakan suara: {voices.Item(1).GetDescription()}")

# Atur kecepatan suara (range: -10 sampai 10, default: 0)
speak.Rate = 1  # Sedikit lebih cepat

# Atur volume (range: 0-100, default: 100)
speak.Volume = 80

torch.multiprocessing.set_start_method('spawn', force=True)

## Setting up torch for gpu utilization
if torch.cuda.is_available():
    torch.backends.cudnn.enabled = True 
    torch.backends.cudnn.benchmark = True
    torch.backends.cudnn.deterministic = True
    torch.set_default_tensor_type('torch.cuda.FloatTensor')

def prep_image(img, inp_dim):
    """
    Prepare image for inputting to the neural network.
    Returns a Variable
    """
    orig_im = img
    dim = orig_im.shape[1], orig_im.shape[0]
    img = (letterbox_image(orig_im, (inp_dim, inp_dim)))
    img_ = img[:, :, ::-1].transpose((2, 0, 1)).copy()
    img_ = torch.from_numpy(img_).float().div(255.0).unsqueeze(0)
    return img_, orig_im, dim

labels = {}
b_boxes = {}
last_speak_time = 0  # Untuk kontrol interval speaking

def write(bboxes, img, classes, colors):
    """
    Draws the bounding box in every frame over the objects that the model detects
    """
    global last_speak_time
    
    class_idx = bboxes
    bboxes = bboxes[1:5]
    bboxes = bboxes.cpu().data.numpy()
    bboxes = bboxes.astype(int)
    b_boxes.update({"bbox": bboxes.tolist()})
    bboxes = torch.from_numpy(bboxes)
    cls = int(class_idx[-1])
    label = "{0}".format(classes[cls])
    labels.update({"Current Object": label})
    color = random.choice(colors)

    ## Put text configuration on frame
    text_str = '%s' % (label) 
    font_face = cv2.FONT_HERSHEY_DUPLEX
    font_scale = 0.6
    font_thickness = 1
    text_w, text_h = cv2.getTextSize(text_str, font_face, font_scale, font_thickness)[0]
    text_pt = (bboxes[0], bboxes[1] - 3)
    text_color = [255, 255, 255]

    ## Distance Measurement for each bounding box
    x, y, w, h = bboxes[0], bboxes[1], bboxes[2], bboxes[3]
    distance = (2 * 3.14 * 180) / (w.item() + h.item() * 360) * 1000 + 3
    
    # Konversi ke centimeter untuk lebih mudah dipahami
    distance_cm = distance * 2.54
    
    feedback = "{} is at {} centimeters".format(
        labels["Current Object"], 
        round(distance_cm)
    )
    
    # Speaking dengan interval (agar tidak terlalu sering berbicara)
    current_time = time.time()
    if current_time - last_speak_time > 3:  # Bicara setiap 3 detik
        speak.Speak(feedback)
        last_speak_time = current_time
    
    print(feedback)
    
    # Tampilkan jarak dalam centimeter di video
    cv2.putText(
        img, 
        str("{:.1f} cm".format(distance_cm)), 
        (int(x), int(y) - 10), 
        cv2.FONT_HERSHEY_DUPLEX, 
        font_scale, 
        (0, 255, 0), 
        font_thickness, 
        cv2.LINE_AA
    )
    
    cv2.rectangle(
        img, 
        (int(bboxes[0]), int(bboxes[1])),
        (int(bboxes[2]), int(bboxes[3])), 
        color, 
        2
    )
    
    cv2.putText(
        img, 
        text_str, 
        (int(text_pt[0]), int(text_pt[1])), 
        font_face, 
        font_scale, 
        text_color, 
        font_thickness, 
        cv2.LINE_AA
    )

    return img

class ObjectDetection:
    def __init__(self, id): 
        print("Initializing camera...")
        self.cap = WebcamVideoStream(src=id).start()
        time.sleep(2.0)  # Tunggu kamera siap
        
        self.cfgfile = "cfg/yolov3.cfg"
        self.weightsfile = "yolov3.weights"
        
        # Jika ingin pakai YOLOv3-tiny (lebih cepat):
        # self.cfgfile = 'cfg/yolov3-tiny.cfg'
        # self.weightsfile = 'yolov3-tiny.weights'
        
        self.confidence = float(0.6)
        self.nms_thesh = float(0.8)
        self.num_classes = 80
        self.classes = load_classes('data/coco.names')
        self.colors = pkl.load(open("pallete", "rb"))
        
        print("Loading YOLO model...")
        self.model = Darknet(self.cfgfile)
        self.CUDA = torch.cuda.is_available()
        self.model.load_weights(self.weightsfile)
        self.model.net_info["height"] = 160
        self.inp_dim = int(self.model.net_info["height"])
        
        # Resolusi video
        self.width = 1280
        self.height = 720
        
        print("Loading network.....")
        if self.CUDA:
            self.model.cuda()
            print("Running on GPU")
        else:
            print("Running on CPU")
            
        print("Network successfully loaded")
        assert self.inp_dim % 32 == 0
        assert self.inp_dim > 32
        self.model.eval()

    def main(self):
        q = queue.Queue()
        
        def frame_render(queue_from_cam):
            frame = self.cap.read()
            if frame is not None:
                frame = cv2.resize(frame, (self.width, self.height))
                queue_from_cam.put(frame)
        
        cam = threading.Thread(target=frame_render, args=(q,))
        cam.start()
        cam.join()
        
        frame = q.get()
        q.task_done()
        
        fps = FPS().start()
        
        try:
            img, orig_im, dim = prep_image(frame, self.inp_dim)
            im_dim = torch.FloatTensor(dim).repeat(1, 2)
            
            if self.CUDA:
                im_dim = im_dim.cuda()
                img = img.cuda()
            
            output = self.model(Variable(img), self.CUDA)
            output = write_results(
                output, 
                self.confidence, 
                self.num_classes, 
                nms=True, 
                nms_conf=self.nms_thesh
            )
            
            if output.numel() > 0:
                output = output.type(torch.half)
                
                if list(output.size()) != [1, 86]:
                    output[:, 1:5] = torch.clamp(
                        output[:, 1:5], 
                        0.0, 
                        float(self.inp_dim)
                    ) / self.inp_dim
                    
                    output[:, [1, 3]] *= frame.shape[1]
                    output[:, [2, 4]] *= frame.shape[0]
                    
                    list(map(
                        lambda boxes: write(boxes, frame, self.classes, self.colors),
                        output
                    ))
        except Exception as e:
            print(f"Error in detection: {e}")
            pass
        
        fps.update()
        fps.stop()
        
        ret, jpeg = cv2.imencode('.jpg', frame)
        
        # Print FPS info (bisa di-comment jika tidak perlu)
        # print("[INFO] elapsed time: {:.2f}".format(fps.elapsed()))
        # print("[INFO] approx. FPS: {:.1f}".format(fps.fps()))

        return jpeg.tobytes()