import cv2

def click_event(event, x, y, flags, params):
    if event == cv2.EVENT_LBUTTONDOWN:
        #print(x, ' ', y)
        font = cv2.FONT_HERSHEY_SIMPLEX
        cv2.putText(img, str(x*3) + ',' + str(y*3), (x, y), font, 1, (255, 0, 0), 2)
        cv2.imshow(new_path, img)
        global mouseX,mouseY
        mouseX = x*3
        mouseY = y*3

def main(path):
    global img,new_path
    new_path = path
    img = cv2.imread(path, 1)
    img = cv2.resize(img, (0,0), fx=0.33, fy=0.33)
    cv2.imshow(path, img)
    cv2.moveWindow(path, 800, 0)
    cv2.setMouseCallback(path, click_event)
    cv2.waitKey(0)
    print(mouseX,mouseY)
    cv2.destroyAllWindows()
    return mouseX,mouseY
