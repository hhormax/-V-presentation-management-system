import cv2
import mediapipe as mp
import time
import win32com.client
Application = win32com.client.Dispatch("PowerPoint.Application")
Presentation = Application.Presentations.Open("D:\\cwPy\\presa.pptx")
Presentation.SlideShowSettings.Run()

mp_hands = mp.solutions.hands
mp_drawing = mp.solutions.drawing_utils
hands = mp_hands.Hands(max_num_hands=1, min_detection_confidence=0.5, min_tracking_confidence=0.5)

last_output_time = 0
output_delay = 1.5
height = 50
first_width = 30
second_width = 300

def process_hands(image, results):
    global last_output_time

    if results.multi_hand_landmarks:
        for hand_landmarks in results.multi_hand_landmarks:
            index_finger_tip = hand_landmarks.landmark[mp_hands.HandLandmark.INDEX_FINGER_TIP]
            index_finger_base = hand_landmarks.landmark[mp_hands.HandLandmark.INDEX_FINGER_MCP]

            index_finger_x = int(index_finger_tip.x * image.shape[1])
            index_finger_y = int(index_finger_tip.y * image.shape[0])
            index_finger_bx = int(index_finger_base.x * image.shape[1])
            index_finger_by = int(index_finger_base.y * image.shape[0])

            if index_finger_y and index_finger_x:
                h = index_finger_by - index_finger_y
                w = index_finger_bx - index_finger_x

                current_time = time.time()

                if abs(h) <= height and -second_width <= w <= -first_width:
                    if current_time - last_output_time >= output_delay:
                        print("Right")
                        Presentation.SlideShowWindow.View.Next()
                        last_output_time = current_time
                elif abs(h) <= height and first_width <= w <= second_width:
                    if current_time - last_output_time >= output_delay:
                        print("Left")
                        Presentation.SlideShowWindow.View.Previous()
                        last_output_time = current_time

cap = cv2.VideoCapture(0)
while cap.isOpened():
    success, image = cap.read()

    if not success:
        print("Ignoring empty camera frame.")
        continue

    image = cv2.cvtColor(cv2.flip(image, 1), cv2.COLOR_BGR2RGB)
    results = hands.process(image)

    process_hands(image, results)

    cv2.imshow('MediaPipe Hands', cv2.cvtColor(image, cv2.COLOR_RGB2BGR))

    key = cv2.waitKey(1)
    if key == 27:
        break

cap.release()
cv2.destroyAllWindows()
