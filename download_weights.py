import urllib.request

print("Downloading yolov3.weights...")
url = "https://pjreddie.com/media/files/yolov3.weights"
urllib.request.urlretrieve(url, "yolov3.weights")
print("Download selesai!")