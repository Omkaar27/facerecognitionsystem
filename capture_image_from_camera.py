from cv2 import*
import cv2
cam_port = 0
cam =cv2. VideoCapture(cam_port)
# reading the input using the camera
if not cam.isOpened():
   print("Error: Could not open camera.")
   exit() 
inp = input('Enter person name')
# If image will detected without any error,
# show result
while True: 
        result,image = cam.read() 
        cv2.imshow(inp, image) 
        if cv2.waitKey(0):
         cv2.imwrite(inp+ ".png", image)
         print("image taken") 

# If captured image is corrupted, moving to else part
else:
        print("No image detected. Please! try again") 
