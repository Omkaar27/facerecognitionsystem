import cv2

cam_port = 0
cam =cv2. VideoCapture(cam_port)

while True:

    inp = input('Enter person name')

    if inp.lower() == "exit":
        break  

    while True:  
        result,image = cam.read() 
        if result:
            # Display the captured image
            cv2.imshow(inp, image)

            # Wait for a key press
            key = cv2.waitKey(1)

            if key == ord('s'):  # Press 's' to save the image
                image_filename = inp + ".png"
                cv2.imwrite(image_filename, image)
                print("Image taken and saved as", image_filename)
                break
            elif key == 27:
                print("No image detected. Please! try again") 
                break
          
    cv2.destroyWindow(inp)
cam.release() 
cv2.destroyAllWindows()
