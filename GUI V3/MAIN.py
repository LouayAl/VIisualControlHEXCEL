# almost final version
# Version 3

from tkinter import *
from tkinter import messagebox
import tkinter
from typing import Self
from PIL import Image, ImageTk, ImageDraw, ImageFont
from PIL.PngImagePlugin import PngInfo  # à verifier si besoin
import cv2
import numpy as np
import pandas as pd
import os
import datetime

# please change the path to your data file (check for several places to change the path in the code below)
data_frame = pd.read_excel(
    "C:\\Users\\louay\\Desktop\\data.xlsx", sheet_name="DATA")

# Create the info window that will open the root window
info = Tk()
info.title("Hexcel Casablanca Inspection Pictures")
info.geometry("370x350")

# Setup camera left and right
left_cam = 0
right_cam = 1

# Tell how many camera
nb_camera_stack = 2
nb_camera_extra = 1

# storing item
itm = None

# for the condition of the take picture button if the analyzing is done the zoom_done will be changed to yes to enable taking pictures
zoom_done = 'no'
# for the condition of the insertion of the item ID and if it's already present in the database or not.
autres = 'no'

# Set-up offset
offset_Lx = 0  # left cam
offset_Ly = 0  # left cam
offset_Rx = 0  # Right cam
offset_Ry = 0  # Right cam

# this function is called after the open root window button is clicked, it creates the root window


def open_root_window():
    # to hide the info window
    info.withdraw()

    # to read from the excel file please note that you need to change the path
    data_frame = pd.read_excel(
        "C:\\Users\\louay\\Desktop\\data.xlsx", sheet_name="DATA")
    data_frame2 = pd.read_excel(
        "C:\\Users\\louay\\Desktop\\AUTRES.xlsx", sheet_name="AUTRES")

    # Tell how many camera
    global nb_camera_stack
    global nb_camera_extra

    # Setup camera left and right
    # left_cam = 0
    # right_cam = 1

    # Set-up offset
    offset_Lx = 405  # left cam
    offset_Ly = 0  # left cam
    offset_Rx = 380  # Right cam
    offset_Ry = 8  # Right cam

    # to activate cameras once the root window is opened

    def camera_activation(string):
        if string == 'yes':
            global nb_camera_extra
            global nb_camera_stack
            nb_camera_stack = 2
            nb_camera_extra = 0
        if string == 'no':
            nb_camera_stack = 0
            nb_camera_extra = 1
        print(nb_camera_stack, nb_camera_extra, '   11111211111')
        return nb_camera_stack, nb_camera_extra

    # once you close the root window it displays the info window so the user can enter another item. note that the function is called in the root.protocol line
    def on_root_close():
        global zoom_done
        cap0.release()
        cap1.release()
        cap2.release()
        root.destroy()
        info.deiconify()

        zoom_done = 'no'

    # Start GUI
    root = Toplevel(info)
    root.geometry("1200x700")
    root.title("Hexcel Casablanca Inspection Pictures")
    root.protocol("WM_DELETE_WINDOW", on_root_close)
    root.state('zoomed')  # Maximize the window

    # add background
    image1 = Image.open("images/background.png")
    test1 = ImageTk.PhotoImage(image1)
    label1 = tkinter.Label(root, image=test1)
    label1.image = test1
    label1.place(x=0, y=0)

    # Title
    label0a = Label(
        root, text="Hexcel Casablanca Inspection Pictures", font="Abadi 16 bold", )
    label0a.grid(row=0, column=1, columnspan=8)

    img = Image.open(r"images/logo hexcel 2.png")
    img = img.resize((int(img.width/2 - 20), int(img.height/2 - 10)))
    img = ImageTk.PhotoImage(img)
    label0b = Label(root)
    label0b.config(image=img)
    label0b.image = img
    label0b.grid(row=0, column=1, columnspan=2, rowspan=1, padx=0, pady=0)

#    adding logo (remember image should be PNG and not JPG)
    img = PhotoImage(file=r"images/logo hexcel 2.png")
    img1 = img.subsample(2, 2)
    label0b = Label(root, image=img1)

    label_cam0 = Label(root)
    label_cam0.frame_num = 0
    label_cam0.grid(row=1, column=1, columnspan=4, rowspan=2, padx=10, pady=2)
    label_cam0.config(borderwidth=0)
    cap0 = cv2.VideoCapture(1, cv2.CAP_DSHOW)
    cap0.set(cv2.CAP_PROP_FRAME_WIDTH, 1920)
    cap0.set(cv2.CAP_PROP_FRAME_HEIGHT, 1024)

    label_cam1 = Label(root)
    label_cam1.frame_num = 0
    label_cam1.grid(row=1, column=0, columnspan=4, rowspan=6, padx=10, pady=2)
    label_cam1.config(borderwidth=0)
    cap1 = cv2.VideoCapture(2, cv2.CAP_DSHOW)
    cap1.set(cv2.CAP_PROP_FRAME_WIDTH, 1920)
    cap1.set(cv2.CAP_PROP_FRAME_HEIGHT, 1024)
    print("Hello cam extra1")

    # label_cam3 = Label(root)
    # label_cam3.frame_num = 0
    # label_cam3.grid(row=1, column=0, columnspan=4, rowspan=6, padx=10, pady=2)
    # label_cam3.config(borderwidth=0)
    # cap3 = cv2.VideoCapture(0, cv2.CAP_DSHOW)

    label_cam2 = Label(root)
    label_cam2.frame_num = 0
    label_cam2.grid(row=1, column=0, columnspan=4, rowspan=6, padx=10, pady=2)
    label_cam2.config(borderwidth=0)
    cap2 = cv2.VideoCapture(1, cv2.CAP_DSHOW)
    cap2.set(cv2.CAP_PROP_FRAME_WIDTH, 1920)
    cap2.set(cv2.CAP_PROP_FRAME_HEIGHT, 1024)
    # adding a line below cameras to not change behavior
    img_separator = PhotoImage(file=r"images/divider.png")
    label_separator = Label(root, image=img_separator)
    label_separator.grid(row=9, column=1, columnspan=15, padx=10, pady=0)

    # #If there are 2 camera, add the second one
    # if nb_camera_stack ==2:
    #     #Camera 1
    #     label_cam0 = Label(root)
    #     label_cam0.frame_num = 0
    #     label_cam0.grid(row=1, column=0, columnspan = 5, rowspan = 6, padx = 10, pady = 0)
    #     cap0= cv2.VideoCapture(0, cv2.CAP_DSHOW)

    if not cap0.isOpened():
        print("Cannot open camera")
        exit()
    if not cap1.isOpened():
        print("Cannot open camera")
        exit()
    if not cap2.isOpened():
        print("Cannot open camera")
        exit()
    # if not cap3.isOpened():
    #     print("Cannot open camera")
    #     exit()

    #     print("Hello cam 2")
    #     label_cam1 = Label(root)
    #     label_cam1.frame_num = 0
    #     label_cam1.grid(row=1, column=6, columnspan = 5, rowspan = 10, padx = 10, pady = 10)
    #     cap1= cv2.VideoCapture(0, cv2.CAP_DSHOW)

    #     if not cap1.isOpened():
    #         nb_camera_stack = nb_camera_stack - 1
    #         print("Cannot open camera. Nb cameras : " + str(nb_camera_stack))
    #         exit()

    # # #If there is extra camera, add it
    # else:
    #     if nb_camera_extra ==1:
    #         print("Hello cam extra")
    #         label_cam2 = Label(root)
    #         label_cam2.frame_num = 0
    #         label_cam2.grid(row=1, column=0, columnspan = 5, rowspan = 10, padx = 10, pady = 10)
    #         cap2= cv2.VideoCapture(0, cv2.CAP_DSHOW)

    # displaying the frames of the video camera

    def show_frames():
        lookup = data_frame[(data_frame["Item"] == itm)]

        if nb_camera_stack == 2 and nb_camera_extra == 0:

            # # Loading the logo image
            # config1 = Image.open("images/config 1.png")
            # IMGwidth, IMGheight = config1.size
            # ratio = (IMGheight - 170) / IMGheight
            # resized_image = config1.resize(
            #     (int(IMGwidth * ratio), IMGheight - 170), Image.LANCZOS)
            # img_config2 = ImageTk.PhotoImage(resized_image)
            # label6b = tkinter.Label(root, image=img_config2)
            # label6b.image = img_config2
            # label6b.place(x=20, y=500)

            # Loading the logo image
            catia1 = Image.open("label images/2.png")
            Cwidth, Cheight = catia1.size
            Cratio = (Cheight - 120) / Cheight
            resized_image = catia1.resize(
                (int(Cwidth * Cratio), Cheight - 120), Image.LANCZOS)
            img_config2 = ImageTk.PhotoImage(resized_image)
            label6b = tkinter.Label(root, image=img_config2)
            label6b.image = img_config2
            label6b.place(x=30, y=615)

            x1 = int(lookup.iloc[0]['cropX1'])
            x2 = int(lookup.iloc[0]['cropX2'])
            y1 = int(lookup.iloc[0]['cropY1'])
            y2 = int(lookup.iloc[0]['cropY2'])

            # print("zoom valuesL:  ", x1, ' ', x2, ' ', y1, ' ', y2, ' ')

            # frame0
            _, frame0 = cap0.read()
            frame0 = cv2.flip(frame0, 1)
            frame0 = frame0[offset_Ly+x1:frame0.shape[0] -
                            offset_Ry-y1, 250 + x2:frame0.shape[1]-offset_Lx]
            frame0_resized = cv2.resize(
                frame0, (0, 0), fx=0.3, fy=0.3, interpolation=cv2.INTER_CUBIC)
            frame0_resized = cv2.putText(
                frame0_resized, 'CAM 0', (5, 10), cv2.FONT_HERSHEY_SIMPLEX, 1/3, (255, 255, 0), 1, cv2.LINE_AA)

            # frame1
            _, frame1 = cap1.read()
            frame1 = cv2.flip(frame1, 1)
            frame1 = frame1[offset_Ry+x1:frame1.shape[0] -
                            offset_Ly-y1, offset_Rx + x2:frame1.shape[1]-250]
            frame1_resized = cv2.resize(
                frame1, (0, 0), fx=0.3, fy=0.3, interpolation=cv2.INTER_AREA)

            frame1_resized = cv2.putText(
                frame1_resized, 'CAM 1', (5, 10), cv2.FONT_HERSHEY_SIMPLEX, 1/3, (255, 255, 0), 1, cv2.LINE_AA)

            # _, frame3 = cap3.read()
            # frame3 = cv2.flip(frame3, 1)
            # frame3 = frame3[offset_Ry:frame3.shape[0] -
            #                 offset_Ly, offset_Rx:frame3.shape[1]]
            # frame3_resized = cv2.resize(
            #     frame3, (0, 0), fx=0.45, fy=0.45, interpolation=cv2.INTER_AREA)
            # frame3_resized = cv2.putText(
            #     frame3_resized, 'CAM 2', (5, 10), cv2.FONT_HERSHEY_SIMPLEX, 1/3, (255, 255, 0), 1, cv2.LINE_AA)

            if left_cam == 0:
                img_stack = np.hstack([frame0, frame1])
                img_stack_resized = np.hstack(
                    [frame0_resized, frame1_resized])
            if left_cam == 1:
                img_stack = np.hstack([frame1, frame0])
                img_stack_resized = np.hstack(
                    [frame1_resized, frame0_resized])

            cv2image_HD = cv2.cvtColor(img_stack, cv2.COLOR_BGR2RGB)
            img_HD = Image.fromarray(cv2image_HD)

            # img_HD = Image.fromarray(hsv)
            imgtk_HD = ImageTk.PhotoImage(image=img_HD)

            cv2image0 = cv2.cvtColor(img_stack_resized, cv2.COLOR_BGR2RGB)

        #     ### CONTOUR ###
        #     ###############
        #     img = cv2image0
        #     gray_img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        #     height, width = gray_img.shape

        #     gray_img = 255 - gray_img
        #     gray_img[gray_img > 90] = 255
        #     gray_img[gray_img <= 90] = 0

        #     kernel = np.ones((3, 3), np.uint8)
        #     gray_img = cv2.morphologyEx(gray_img, cv2.MORPH_CLOSE, kernel,(-1,-1),iterations = 2)

        #     gray_img = np.uint8(gray_img)
        #     canny_img = cv2.Canny(gray_img, 100, 200, 3)

        #     resized_gray = cv2.resize(gray_img, (500, 250))
        #     resized_canny = cv2.resize(canny_img, (800, 400))

        #     minLineLength = 500
        #     maxLineGap = 100
        #     lines = cv2.HoughLinesP(canny_img, 1, np.pi / 180, 8, None, 20, 30)

        #     # Remove lines that are not horizontal or vertical
        #     horizontal_lines = []
        #     vertical_lines = []
        #     for line in lines:
        #         x1, y1, x2, y2 = line[0]
        #         if abs(y2 - y1) < 7: # horizontal line
        #             horizontal_lines.append(line)
        #         elif abs(x2 - x1) < 9.5: # vertical line
        #             vertical_lines.append(line)
        #     # Draw horizontal lines
        #     for line in horizontal_lines:
        #         x1, y1, x2, y2 = line[0]
        #         cv2.line(img, (x1, y1), (x2, y2), (0, 0, 0), 2)
        #     # Draw vertical lines
        #     for line in vertical_lines:
        #         x1, y1, x2, y2 = line[0]
        #         cv2.line(img, (x1, y1), (x2, y2), (0, 0, 0), 2)

        #     # Slope
        #     for line in lines:
        #         for x1, y1, x2, y2 in line:
        #             slope = (y2 - y1) / (x2 - x1)
        #             line = [x1, y1, x2, y2, slope]
        #             #print(line)
        #             #cv2.line(img, (x1, y1), (x2, y2), (0, 0, 255), 2)

        #     #draw all lines 1
        #     for line in lines:
        #         for x1,y1,x2,y2 in line:
        #             cv2.line(img,(x1,y1),(x2,y2),(0,255,255),1)

        #     all_lines_x_sorted = sorted(vertical_lines, key=lambda k: k[0][0])

        #     for x1,y1,x2,y2 in all_lines_x_sorted[0]:
        #         cv2.line(img,(x1,y1),(x2,y2),(0,0,255),2)
        #         x0_x=x1
        #         x0_y=y1

        #     all_lines_x_sorted = sorted(vertical_lines, key=lambda k: -k[0][0])
        #     for x1,y1,x2,y2 in all_lines_x_sorted[0]:
        #         cv2.line(img,(x1,y1),(x2,y2),(0,0,255),2)
        #         x1_x=x1
        #         x1_y=y1

        #     all_lines_y_sorted = sorted(horizontal_lines, key=lambda k: k[0][1])

        #     for x1,y1,x2,y2 in all_lines_y_sorted[0]:
        #         cv2.line(img,(x1,y1),(x2,y2),(0,0,255),2)
        #         y0_x=x1
        #         y0_y=y1

        #     all_lines_y_sorted = sorted(horizontal_lines, key=lambda k: -k[0][1])
        #     for x1,y1,x2,y2 in all_lines_y_sorted[0]:
        #         cv2.line(img,(x1,y1),(x2,y2),(0,0,255),2)
        #         y1_x=x1
        #         y1_y=y1

        #    ### POTTING ###
        #     ###############

        #     if  (lookup.iloc[0]['Potting']) == 'yes':

        #         # Convert the image to HSV color space
        #         hsv = cv2.cvtColor(cv2image0, cv2.COLOR_BGR2HSV)

        #         # Define the lower and upper bounds of the white color in HSV
        #         white_lower = np.array([0, 0, 0])
        #         white_upper = np.array([127, 95, 86])

        #         # Create a mask that only selects the white color in the image
        #         mask = cv2.inRange(hsv, white_lower, white_upper)

        #         kernel = np.ones((5, 5), np.uint8)

        #         mask = cv2.erode(mask, kernel, iterations=2)
        #         mask = cv2.dilate(mask, kernel, iterations=3)
        #         contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        #         # Find the largest contour that matches the color
        #         if contours:
        #             largest_contour = max(contours, key=cv2.contourArea)
        #         else:
        #             messagebox.showinfo("PIECE NOT IN POSITION", "no contours detected, Please press retour and place your piece correctly on the table then run the application again")

        #         # print(largest_contour)

        #         # Draw the largest contour on the original image
        #         cv2.drawContours(cv2image0, [largest_contour],-1,(0, 255, 0), 2)

        #         # cv2.imshow('Built-in Camera', cv2image_HD)

        #         # Draw the enclosing circle
        #         (pot_x,pot_y),radius = cv2.minEnclosingCircle(largest_contour)
        #         center_potting = (int(pot_x),int(pot_y))
        #         radius_potting = int(radius)
        #         cv2.circle(cv2image0,center_potting,radius_potting,(0,0,255),1)

        #         font = cv2.FONT_HERSHEY_SIMPLEX
        #         T= "Potting : r=" + str(radius_potting) + " (x,y)=" + str(center_potting)
        #         cv2.putText(cv2image0, T ,center_potting, font, 1/3,(0,0,255),1,cv2.LINE_AA)

        #         #cv2.arrowedLine(img,center_potting,(x0_x,int(pot_y)),(0,255,0),1,tipLength = 1/50)
        #         #cv2.arrowedLine(img,(int(pot_x),y0_y),center_potting,(0,255,0),1,tipLength = 1/50)

        #         dist_x0 = abs(center_potting[0] - x0_x)
        #         dist_x1 = abs(center_potting[0] - x1_x)
        #         dist_y0 = abs(center_potting[1] - y0_y)
        #         dist_y1 = abs(center_potting[1] - y1_y)

        #         if dist_x0 < dist_x1 :
        #             cv2.arrowedLine(cv2image0,center_potting,(x0_x,int(pot_y)),(0,255,0),1,tipLength = 1/50)
        #             cv2.putText(cv2image0, "Distance_x: {:.2f}".format(dist_x0), (center_potting[0]-130, center_potting[1]+20),
        #                         cv2.FONT_HERSHEY_SIMPLEX, 1/3, (255, 0, 0), 1)
        #         else:
        #             cv2.arrowedLine(cv2image0,center_potting,(x1_x,int(pot_y)),(0,255,0),1,tipLength = 1/50)
        #             cv2.putText(cv2image0, "Distance_x: {:.2f}".format(dist_x1), (center_potting[0]-130, center_potting[1]+20),
        #                         cv2.FONT_HERSHEY_SIMPLEX, 1/3, (255,0,0), 1)

        #         if dist_y0 < dist_y1 :
        #             cv2.arrowedLine(cv2image0,center_potting,(int(pot_x),y0_y),(0,255,0),1,tipLength = 1/50)
        #             cv2.putText(cv2image0, "Distance_y: {:.2f}".format(dist_y0), (center_potting[0]+10, center_potting[1]-10),
        #                         cv2.FONT_HERSHEY_SIMPLEX, 1/3, (255, 0,0), 1)
        #         else:
        #             cv2.arrowedLine(cv2image0,center_potting,(int(pot_x),y1_y),(0,255,0),1,tipLength = 1/50)
        #             cv2.putText(cv2image0, "Distance_y: {:.2f}".format(dist_y1), (center_potting[0]+10, center_potting[1]-10),
        #                         cv2.FONT_HERSHEY_SIMPLEX, 1/3, (255, 0,0), 1)

        #     ### SPLICE ###
        #     ##############

        #     # Calculer les valeurs minimales et maximales pour chaque canal HSV
        #     h_min = min(24, 30, 34, 37) - 120
        #     h_max = max(24, 30, 34, 37) + 120
        #     s_min = min(0.38, 0.50, 0.73, 0.72) - 0.1
        #     s_max = max(0.38, 0.50, 0.73, 0.72) + 0.1
        #     v_min = min(0.63, 0.63, 0.63, 0.69) - 0.1
        #     v_max = max(0.63, 0.63, 0.63, 0.69) + 0.1

        #     # Définir la plage de couleurs en fonction des valeurs minimales et maximales pour chaque canal HSV
        #     lower_color = np.array([h_min, s_min * 255, v_min * 255])
        #     upper_color = np.array([h_max, s_max * 255, v_max * 255])
        #     hsv = cv2.cvtColor(cv2image0, cv2.COLOR_BGR2HSV)

        #     # Create a mask that only selects the beige color in the image
        #     mask = cv2.inRange(hsv, lower_color, upper_color)
        #     # Find the contours in the mask
        #     contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        #     # Sort the contours by area in descending order and select the first three contours
        #     contours = sorted(contours, key=cv2.contourArea, reverse=True)[:5]

        #     # Draw the contours on the original image
        #     cv2.drawContours(img, contours, -1, (255, 0, 0), 2)

        #     # ### PRINT TEXT IN IMG ###
        #     # #########################
        #     # font = cv2.FONT_HERSHEY_SIMPLEX
        #     # T= "X0 = (" + str(x0_x) + ","+ str(x0_y) +")"
        #     # cv2.putText(img, T ,(x0_x+5,x0_y), font, 1/3,(0,0,255),1,cv2.LINE_AA)

        #     # T= "Y0 = (" + str(x0_x) + ","+ str(x0_y) +")"
        #     # cv2.putText(img, T ,(y0_x-35,y0_y-10), font, 1/3,(0,0,255),1,cv2.LINE_AA)

        #     # T = 'Potting: ' + str(item_details.iloc[0]['Potting'])
        #     # cv2.putText(img, T,(5,20), font, 1/3,(255,0,0),1,cv2.LINE_AA)
        #     # T = 'Item: ' + str(item_details.iloc[0]['Item'])
        #     # cv2.putText(img, T,(5,35), font, 1/3,(255,0,0),1,cv2.LINE_AA)

        #     # T = 'Potting: ' + str(item_details.iloc[0]['Pot_D']) + "_" + str(item_details.iloc[0]['Pot_D']) + "_" + str(item_details.iloc[0]['Pot_D'])

        #     # cv2.putText(img, T,(5,50), font, 1/3,(255,0,0),1,cv2.LINE_AA)

            img = Image.fromarray(cv2image0)
            imgtk = ImageTk.PhotoImage(image=img)
            label_cam0.imgtk = imgtk
            label_cam0.frame_num += 1.
            label_cam0.configure(image=imgtk)
            label_cam0.after(100, show_frames)

        if nb_camera_stack == 0:

            config2 = Image.open("images/config 2.png")
            IMGwidth, IMGheight = config2.size
            ratio = (IMGheight - 170) / IMGheight
            resized_image = config2.resize(
                (int(IMGwidth * ratio), IMGheight - 170), Image.LANCZOS)
            img_config2 = ImageTk.PhotoImage(resized_image)
            label6b = tkinter.Label(root, image=img_config2)
            label6b.image = img_config2

            label6b.place(x=20, y=450)

            # Loading the logo image
            catia1 = Image.open("label images/2.png")
            Cwidth, Cheight = catia1.size
            Cratio = (Cheight - 120) / Cheight
            resized_image = catia1.resize(
                (int(Cwidth * Cratio), Cheight - 120), Image.LANCZOS)
            img_config2 = ImageTk.PhotoImage(resized_image)
            label6b = tkinter.Label(root, image=img_config2)
            label6b.image = img_config2
            label6b.place(x=30, y=600)

            _, frame2 = cap2.read()
            # frame2 = cv2.flip(frame2, 1)
            frame2 = frame2[offset_Ry:frame2.shape[0] -
                            offset_Ly, 0:frame2.shape[1]-offset_Lx]
            frame2_resized = cv2.resize(
                frame2, (0, 0), fx=0.35, fy=0.35, interpolation=cv2.INTER_AREA)
            frame2_resized = cv2.putText(frame2_resized, 'CAM EXTRA 1', (
                5, 10), cv2.FONT_HERSHEY_SIMPLEX, 1/3, (255, 255, 0), 1, cv2.LINE_AA)

            if left_cam == 0:
                img_stack = np.hstack([frame2])
                img_stack_resized = np.hstack([frame2_resized])
            if left_cam == 1:
                img_stack = np.hstack([frame2, frame2])
                img_stack_resized = np.hstack([frame2_resized])

            cv2image_HD = cv2.cvtColor(img_stack, cv2.COLOR_BGR2RGB)
            img_HD = Image.fromarray(cv2image_HD)
            imgtk_HD = ImageTk.PhotoImage(image=img_HD)

            cv2image0 = cv2.cvtColor(img_stack_resized, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(cv2image0)
            imgtk = ImageTk.PhotoImage(image=img)
            label_cam2.imgtk = imgtk
            label_cam2.frame_num += 1.
            label_cam2.configure(image=imgtk)
            label_cam2.after(100, show_frames)

    #  this function is not used
    def callback_item2program(sv):

        lookup = data_frame[(data_frame["Item"] == sv.get())]
        if lookup.shape[0] == 1:
            entry3.delete(0, END)  # deletes the current value
            entry3.insert(0, lookup.iloc[0]['Config'])
            entry4.delete(0, END)  # deletes the current value
            entry4.insert(0, lookup.iloc[0]['Potting'])
            camera_activation(lookup.iloc[0]['cam1'])
            print(lookup['cam1'].values[0], '111111111')

        else:
            entry3.delete(0, END)  # deletes the current value
            entry4.delete(0, END)  # deletes the current value
        # camera_activation(lookup['cam1'].values[0])
    # sv = StringVar()
    # sv.set(itm)
    # # camera_activation(sv.get())
    # sv.trace("w", lambda name, index, mode, sv=sv: callback_item2program(sv))

    def key_pressed(event):
        take_pic()
    # taking a picture and saving it to images gui folder beneath the specific batch folder

    def take_pic():
        global zoom_done

        # Create a DataFrame with the data you want to insert
        data = {'Item': [itm],
                'Customer': [customer],
                'Batch': [batch],
                'Config': [config],
                'Defauts': ["N/A"]}
        df = pd.DataFrame(data)

        # Load the existing Excel file
        excel_file = 'C:/Users/louay/Desktop/records.xlsx'
        sheet_name = 'RECORDS'
        existing_data = pd.read_excel(excel_file, sheet_name=sheet_name)

        # Reset the index of the existing_data DataFrame
        existing_data.reset_index(drop=True, inplace=True)

        # Append the new data to the existing data
        new_data = pd.concat([existing_data, df], ignore_index=True)

        # Save the updated data back to the Excel file
        new_data.to_excel(excel_file, sheet_name=sheet_name, index=False)

        lookup = data_frame[(data_frame["Item"] == itm)]
        Longueur = lookup.iloc[0]['Longueur']
        Largeur = lookup.iloc[0]['Largeur']
        Nb_parties = lookup.iloc[0]['Nb parties']
        Potting = lookup.iloc[0]['Potting']
        Cx = lookup.iloc[0]['Cx']
        Cy = lookup.iloc[0]['Cy']
        R = lookup.iloc[0]['R']

        hex_img = Image.open("images/logo hexcel1.png")
        hexcel_titre = "Hexcel Vision Control"
        hex_resized = hex_img.resize((71, 25))

        if zoom_done == 'yes' and (lookup.iloc[0]['cam1']) == 'yes':
            print(imgtk2, 'zoom1')
            w = imgtk2.width()
            h = imgtk2.height()
            merged_image = Image.new("RGB", (w+400, h + 400), color="white")

            # add background
            image1 = Image.open("images/background.png")
            # Paste the zoomed images onto the merged image
            merged_image.paste(image1, (0, 0))
            merged_image.paste(hex_resized, (20, 10))
            merged_image.paste(ImageTk.getimage(imgtk2), (0, 40))
            merged_image.paste(ImageTk.getimage(imgtk_zoomed1), (0, h+40))
            merged_image.paste(ImageTk.getimage(imgtk_zoomed2), (200, h+40))
            merged_image.paste(ImageTk.getimage(imgtk_zoomed3), (400, h+40))
            merged_image.paste(ImageTk.getimage(imgtk_zoomed4), (600, h+40))
            merged_image.paste(ImageTk.getimage(imgtk_zoomed5), (800, h+40))

            title = '/' + itm + '_' + batch + '.png'
            infos = 'Longueur: ' + str(Longueur) + ' Largeur: ' + str(Largeur) + ' Nb_parties:' + str(
                Nb_parties) + ' Potting: ' + str(Potting) + ' Cx: ' + str(Cx) + ' Cy: ' + str(Cy) + ' R: ' + str(R)
            date_string = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            draw = ImageDraw.Draw(merged_image)
            font = ImageFont.truetype("arial.ttf", 16)
            # font2 = ImageFont.truetype("arial.ttf", 16,weight='bold')

            # change values so you can put the text wherever you want
            draw.text((100, 12), hexcel_titre, (0, 0, 0), font)
            draw.text((10, h + 190), infos, (0, 0, 0), font)
            draw.text((10, h + 250), title, (0, 0, 0), font)
            draw.text((180, h + 250), date_string, (0, 0, 0), font)

            merged_image.save("Images GUI/" + customer + title)

            print("Zoomed images saved1.")
            # insert_excel(lookup)
            messagebox.showinfo("Opération terminée", "Photo enregistrée")

        if zoom_done == 'yes' and (lookup.iloc[0]['cam1']) == 'no':

            print(imgtk2, 'zoom2')
            w = imgtk2.width()
            h = imgtk2.height()
            merged_image = Image.new("RGB", (w+77, h + 230), color="white")
            # add background
            image1 = Image.open("images/background.png")
            # Paste the zoomed images onto the merged image
            merged_image.paste(image1, (0, 0))
            merged_image.paste(hex_resized, (20, 10))
            merged_image.paste(ImageTk.getimage(imgtk2), (0, 40))
            merged_image.paste(ImageTk.getimage(imgtk_zoomed2), (0, h+40))
            merged_image.paste(ImageTk.getimage(imgtk_zoomed3), (150, h+40))
            merged_image.paste(ImageTk.getimage(imgtk_zoomed4), (300, h+40))

            title = '/' + itm + '_' + batch + '.png'
            infos = 'Longueur: ' + str(Longueur) + ' Largeur: ' + str(
                Largeur) + ' Nb_parties:' + str(Nb_parties) + ' Potting: ' + str(Potting)
            date_string = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            draw = ImageDraw.Draw(merged_image)
            font = ImageFont.truetype("arial.ttf", 16)
            draw.text((100, 12), hexcel_titre, (0, 0, 0), font)
            draw.text((10, h+190), infos, (0, 0, 0), font)
            draw.text((10, h + 210), title, (0, 0, 0), font)
            draw.text((150, h + 210), date_string, (0, 0, 0), font)

            merged_image.save("Images GUI/" + customer + title)

            print("Zoomed images saved2.")
            messagebox.showinfo("Opération terminée", "Photo enregistrée")

        if not zoom_done == 'yes':
            messagebox.showwarning(
                "Image n'est pas analysé", "Merci d'appuyer sur Analyse avant de prendre une Image")
            return

        zoom_done = 'no'

    # function to save an zoom region as an image to be used in the analyzing part
    def zoom(img2, x, y):
        # Define the coordinates of the region to zoom in on
        width = 60  # Width of the region
        height = 60  # Height of the region
        # Crop the region of interest
        region = img2.crop((x, y, x + width, y + height))
        # Define the zoom factor
        zoom_factor = 3
        # Calculate the new size of the zoomed region
        new_width = width * zoom_factor - 10
        new_height = height * zoom_factor - 10
        # Resize the zoomed region
        zoomed_region = region.resize((new_width, new_height), Image.LANCZOS)
        # Create a copy of the resized image to apply the zoomed region
        zoomed_image = img2
        # Paste the zoomed region back into the zoomed image
        # zoomed_image.paste(zoomed_region, (x, y, x + new_width, y + new_height))
        zoomed_image = zoomed_region
        # Convert the zoomed image to PhotoImage
        imgtk_zoomed = ImageTk.PhotoImage(zoomed_image)
        return imgtk_zoomed

    # takes a picture of and display it as an image combined with the zoomed regions
    def Analyze():
        global imgtk2, imgtk_zoomed1, imgtk_zoomed2, imgtk_zoomed3, imgtk_zoomed4, imgtk_zoomed5, zoom_done
        lookup = data_frame[(data_frame["Item"] == itm)]

        if (lookup.iloc[0]['cam1']) == 'yes':
            # Area for analysis
            label_workspace1 = Label(root)
            label_workspace1.place(x=510, y=268)
            label_workspace2 = Label(root)
            label_workspace2.place(x=890, y=268)

            # Area for zoom1
            label_zoom1 = Label(root)
            label_zoom1.place(x=510, y=551)
            # Area for zoom2
            label_zoom2 = Label(root)
            label_zoom2.place(x=585+95, y=551)
            # Area for zoom3
            label_zoom3 = Label(root)
            label_zoom3.place(x=740+95, y=551)
            # Area for zoom4
            label_zoom4 = Label(root)
            label_zoom4.place(x=895+95, y=551)
            # Area for zoom5
            label_zoom5 = Label(root)
            label_zoom5.place(x=1050+95, y=551)

            # frame0
            _, frame0 = cap0.read()
            frame0 = cv2.flip(frame0, 1)
            x1 = int(lookup.iloc[0]['cropX1'])
            x2 = int(lookup.iloc[0]['cropX2'])
            y1 = int(lookup.iloc[0]['cropY1'])
            y2 = int(lookup.iloc[0]['cropY2'])

            print("zoom valuesL:  ", x1, ' ', x2, ' ', y1, ' ', y2, ' ')
            frame0 = frame0[offset_Ly+x1:frame0.shape[0] -
                            offset_Ry-y1, 250:frame0.shape[1]-offset_Lx]
            frame0_resized = cv2.resize(
                frame0, (0, 0), fx=0.3, fy=0.3, interpolation=cv2.INTER_CUBIC)
            frame0_resized = cv2.putText(
                frame0_resized, 'CAM 0', (5, 10), cv2.FONT_HERSHEY_SIMPLEX, 1/3, (255, 255, 0), 1, cv2.LINE_AA)

            # frame1
            _, frame1 = cap1.read()
            frame1 = cv2.flip(frame1, 1)
            frame1 = frame1[offset_Ry+x1:frame1.shape[0] -
                            offset_Ly-y1, offset_Rx:frame1.shape[1]-250]
            frame1_resized = cv2.resize(
                frame1, (0, 0), fx=0.3, fy=0.3, interpolation=cv2.INTER_AREA)

            frame1_resized = cv2.putText(
                frame1_resized, 'CAM 1', (5, 10), cv2.FONT_HERSHEY_SIMPLEX, 1/3, (255, 255, 0), 1, cv2.LINE_AA)

            # frame0 = frame0[offset_Ly:frame0.shape[0] -
            # offset_Ry, 0:frame0.shape[1]-offset_Lx]
            # frame1 = frame1[offset_Ry:frame1.shape[0] -
            # offset_Ly, offset_Rx:frame1.shape[1]]
            img_stack = np.hstack([frame0_resized, frame1_resized])

            img_stack_HD = np.hstack([frame0, frame1])
            IMGheight, IMGwidth, _ = img_stack_HD.shape
            ratio = (IMGheight - 800) / IMGheight
# provisoire !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            # resized_image = cv2.resize(img_stack_HD, (int(
            #     IMGwidth * ratio), IMGheight - 800), interpolation=cv2.INTER_AREA)
            cv2image3 = cv2.cvtColor(img_stack, cv2.COLOR_BGR2RGB)

            ### CONTOUR ###
            ###############
            img = cv2image3
            gray_img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            height, width = gray_img.shape

            gray_img = 255 - gray_img
            gray_img[gray_img > 90] = 255
            gray_img[gray_img <= 90] = 0

            kernel = np.ones((3, 3), np.uint8)
            gray_img = cv2.morphologyEx(
                gray_img, cv2.MORPH_CLOSE, kernel, (-1, -1), iterations=2)

            gray_img = np.uint8(gray_img)
            canny_img = cv2.Canny(gray_img, 100, 200, 3)

            resized_gray = cv2.resize(gray_img, (500, 250))
            resized_canny = cv2.resize(canny_img, (800, 400))

            minLineLength = 500
            maxLineGap = 100
            lines = cv2.HoughLinesP(canny_img, 1, np.pi / 180, 8, None, 20, 30)

            # Remove lines that are not horizontal or vertical
            horizontal_lines = []
            vertical_lines = []
            for line in lines:
                x1, y1, x2, y2 = line[0]
                if abs(y2 - y1) < 7:  # horizontal line
                    horizontal_lines.append(line)
                elif abs(x2 - x1) < 9.5:  # vertical line
                    vertical_lines.append(line)
            # Draw horizontal lines
            for line in horizontal_lines:
                x1, y1, x2, y2 = line[0]
                cv2.line(img, (x1, y1), (x2, y2), (0, 0, 0), 2)
            # Draw vertical lines
            for line in vertical_lines:
                x1, y1, x2, y2 = line[0]
                cv2.line(img, (x1, y1), (x2, y2), (0, 0, 0), 2)

            # Slope
            for line in lines:
                for x1, y1, x2, y2 in line:
                    slope = (y2 - y1) / (x2 - x1)
                    line = [x1, y1, x2, y2, slope]
                    # print(line)
                    # cv2.line(img, (x1, y1), (x2, y2), (0, 0, 255), 2)

            # draw all lines 1
            for line in lines:
                for x1, y1, x2, y2 in line:
                    cv2.line(img, (x1, y1), (x2, y2), (0, 255, 255), 1)

            all_lines_x_sorted = sorted(vertical_lines, key=lambda k: k[0][0])

            for x1, y1, x2, y2 in all_lines_x_sorted[0]:
                cv2.line(img, (x1, y1), (x2, y2), (0, 0, 255), 2)
                x0_x = x1
                x0_y = y1

            all_lines_x_sorted = sorted(vertical_lines, key=lambda k: -k[0][0])
            for x1, y1, x2, y2 in all_lines_x_sorted[0]:
                cv2.line(img, (x1, y1), (x2, y2), (0, 0, 255), 2)
                x1_x = x1
                x1_y = y1

            all_lines_y_sorted = sorted(
                horizontal_lines, key=lambda k: k[0][1])

            for x1, y1, x2, y2 in all_lines_y_sorted[0]:
                cv2.line(img, (x1, y1), (x2, y2), (0, 0, 255), 2)
                y0_x = x1
                y0_y = y1

            all_lines_y_sorted = sorted(
                horizontal_lines, key=lambda k: -k[0][1])
            for x1, y1, x2, y2 in all_lines_y_sorted[0]:
                cv2.line(img, (x1, y1), (x2, y2), (0, 0, 255), 2)
                y1_x = x1
                y1_y = y1

           ### POTTING ###
            ###############

            if (lookup.iloc[0]['Potting']) == 'yes':

                # Convert the image to HSV color space
                hsv = cv2.cvtColor(cv2image3, cv2.COLOR_BGR2HSV)

                # Define the lower and upper bounds of the white color in HSV
                white_lower = np.array([0, 0, 0])
                white_upper = np.array([127, 95, 86])

                # Create a mask that only selects the white color in the image
                mask = cv2.inRange(hsv, white_lower, white_upper)

                kernel = np.ones((5, 5), np.uint8)

                mask = cv2.erode(mask, kernel, iterations=2)
                mask = cv2.dilate(mask, kernel, iterations=3)
                contours, _ = cv2.findContours(
                    mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

                # Find the largest contour that matches the color
                if contours:
                    largest_contour = max(contours, key=cv2.contourArea)
                else:
                    messagebox.showinfo(
                        "PIECE NOT IN POSITION", "no contours detected, Please press retour and place your piece correctly on the table then run the application again")
                    Retour()

                # print(largest_contour)

                # Draw the largest contour on the original image
                cv2.drawContours(
                    cv2image3, [largest_contour], -1, (0, 255, 0), 2)

                # cv2.imshow('Built-in Camera', cv2image_HD)

                # Draw the enclosing circle
                (pot_x, pot_y), radius = cv2.minEnclosingCircle(largest_contour)
                center_potting = (int(pot_x), int(pot_y))
                radius_potting = int(radius)
                cv2.circle(cv2image3, center_potting,
                           radius_potting, (0, 0, 255), 1)

                font = cv2.FONT_HERSHEY_SIMPLEX
                T = "Potting : r=" + \
                    str(radius_potting) + " (x,y)=" + str(center_potting)
                cv2.putText(cv2image3, T, center_potting, font,
                            1/3, (0, 0, 255), 1, cv2.LINE_AA)

                # cv2.arrowedLine(img,center_potting,(x0_x,int(pot_y)),(0,255,0),1,tipLength = 1/50)
                # cv2.arrowedLine(img,(int(pot_x),y0_y),center_potting,(0,255,0),1,tipLength = 1/50)

                dist_x0 = abs(center_potting[0] - x0_x)
                dist_x1 = abs(center_potting[0] - x1_x)
                dist_y0 = abs(center_potting[1] - y0_y)
                dist_y1 = abs(center_potting[1] - y1_y)

                if dist_x0 < dist_x1:
                    cv2.arrowedLine(cv2image3, center_potting, (x0_x, int(
                        pot_y)), (0, 255, 0), 1, tipLength=1/50)
                    cv2.putText(cv2image3, "Distance_x: {:.2f}".format(dist_x0), (center_potting[0]-130, center_potting[1]+20),
                                cv2.FONT_HERSHEY_SIMPLEX, 1/3, (255, 0, 0), 1)
                else:
                    cv2.arrowedLine(cv2image3, center_potting, (x1_x, int(
                        pot_y)), (0, 255, 0), 1, tipLength=1/50)
                    cv2.putText(cv2image3, "Distance_x: {:.2f}".format(dist_x1), (center_potting[0]-130, center_potting[1]+20),
                                cv2.FONT_HERSHEY_SIMPLEX, 1/3, (255, 0, 0), 1)

                if dist_y0 < dist_y1:
                    cv2.arrowedLine(cv2image3, center_potting, (int(
                        pot_x), y0_y), (0, 255, 0), 1, tipLength=1/50)
                    cv2.putText(cv2image3, "Distance_y: {:.2f}".format(dist_y0), (center_potting[0]+10, center_potting[1]-10),
                                cv2.FONT_HERSHEY_SIMPLEX, 1/3, (255, 0, 0), 1)
                else:
                    cv2.arrowedLine(cv2image3, center_potting, (int(
                        pot_x), y1_y), (0, 255, 0), 1, tipLength=1/50)
                    cv2.putText(cv2image3, "Distance_y: {:.2f}".format(dist_y1), (center_potting[0]+10, center_potting[1]-10),
                                cv2.FONT_HERSHEY_SIMPLEX, 1/3, (255, 0, 0), 1)

            ### SPLICE ###
            ##############

            # Calculer les valeurs minimales et maximales pour chaque canal HSV
            h_min = min(24, 30, 34, 37) - 120
            h_max = max(24, 30, 34, 37) + 120
            s_min = min(0.38, 0.50, 0.73, 0.72) - 0.1
            s_max = max(0.38, 0.50, 0.73, 0.72) + 0.1
            v_min = min(0.63, 0.63, 0.63, 0.69) - 0.1
            v_max = max(0.63, 0.63, 0.63, 0.69) + 0.1

            # Définir la plage de couleurs en fonction des valeurs minimales et maximales pour chaque canal HSV
            lower_color = np.array([h_min, s_min * 255, v_min * 255])
            upper_color = np.array([h_max, s_max * 255, v_max * 255])
            hsv = cv2.cvtColor(cv2image3, cv2.COLOR_BGR2HSV)

            # Create a mask that only selects the beige color in the image
            mask = cv2.inRange(hsv, lower_color, upper_color)
            # Find the contours in the mask
            contours, _ = cv2.findContours(
                mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

            # Sort the contours by area in descending order and select the first three contours
            contours = sorted(contours, key=cv2.contourArea, reverse=True)[:5]

            # Draw the contours on the original image
            cv2.drawContours(img, contours, -1, (255, 0, 0), 2)

            # ### PRINT TEXT IN IMG ###
            # #########################
            # font = cv2.FONT_HERSHEY_SIMPLEX
            # T= "X0 = (" + str(x0_x) + ","+ str(x0_y) +")"
            # cv2.putText(img, T ,(x0_x+5,x0_y), font, 1/3,(0,0,255),1,cv2.LINE_AA)

            # T= "Y0 = (" + str(x0_x) + ","+ str(x0_y) +")"
            # cv2.putText(img, T ,(y0_x-35,y0_y-10), font, 1/3,(0,0,255),1,cv2.LINE_AA)

            # T = 'Potting: ' + str(item_details.iloc[0]['Potting'])
            # cv2.putText(img, T,(5,20), font, 1/3,(255,0,0),1,cv2.LINE_AA)
            # T = 'Item: ' + str(item_details.iloc[0]['Item'])
            # cv2.putText(img, T,(5,35), font, 1/3,(255,0,0),1,cv2.LINE_AA)

            # T = 'Potting: ' + str(item_details.iloc[0]['Pot_D']) + "_" + str(item_details.iloc[0]['Pot_D']) + "_" + str(item_details.iloc[0]['Pot_D'])

            # cv2.putText(img, T,(5,50), font, 1/3,(255,0,0),1,cv2.LINE_AA)

            img2 = Image.fromarray(cv2image3)
            imgtk2 = ImageTk.PhotoImage(image=img2)
            label_workspace1.imgtk = imgtk2
            label_workspace1.configure(image=imgtk2)

            # img3 = Image.fromarray(frame1_resized)
            # imgtk32 = ImageTk.PhotoImage(image=img3)
            # label_workspace2.imgtk = imgtk32
            # label_workspace2.configure(image=imgtk32)

            X1 = (lookup.iloc[0]['Zoom1x'])
            Y1 = (lookup.iloc[0]['Zoom1y'])

            X2 = (lookup.iloc[0]['Zoom2x'])
            Y2 = (lookup.iloc[0]['Zoom2y'])

            X3 = (lookup.iloc[0]['Zoom3x'])
            Y3 = (lookup.iloc[0]['Zoom3y'])

            X4 = (lookup.iloc[0]['Zoom4x'])
            Y4 = (lookup.iloc[0]['Zoom4y'])

            X5 = (lookup.iloc[0]['Zoom5x'])
            Y5 = (lookup.iloc[0]['Zoom5y'])

            imgtk_zoomed1 = zoom(img2, X1, Y1)
            imgtk_zoomed2 = zoom(img2, X2, Y2)
            imgtk_zoomed3 = zoom(img2, X3, Y3)
            imgtk_zoomed4 = zoom(img2, X4, Y4)
            imgtk_zoomed5 = zoom(img2, X5, Y5)

            # Display the zoomed image1 in the label_zoom1
            label_zoom1.imgtk = imgtk_zoomed1
            label_zoom1.configure(image=imgtk_zoomed1)
            # Display the zoomed image2 in the label_zoom2
            label_zoom2.imgtk = imgtk_zoomed2
            label_zoom2.configure(image=imgtk_zoomed2)
            # Display the zoomed image3 in the label_zoom3
            label_zoom3.imgtk = imgtk_zoomed3
            label_zoom3.configure(image=imgtk_zoomed3)
            # Display the zoomed image4 in the label_zoom4
            label_zoom4.imgtk = imgtk_zoomed4
            label_zoom4.configure(image=imgtk_zoomed4)
            # Display the zoomed image5 in the label_zoom5
            label_zoom5.imgtk = imgtk_zoomed5
            label_zoom5.configure(image=imgtk_zoomed5)

            zoom_done = 'yes'

        if (lookup.iloc[0]['cam1']) == 'no':
            # #Area for analysis
            label_workspace = Label(root)
            label_workspace.place(x=702, y=268)
            # Area for zoom1
            # label_zoom1= Label(root)
            # label_zoom1.place(x=550,y=562)
            # Area for zoom2
            label_zoom2 = Label(root)
            label_zoom2.place(x=750, y=170)
            # Area for zoom3
            label_zoom3 = Label(root)
            label_zoom3.place(x=550, y=400)
            # Area for zoom4
            label_zoom4 = Label(root)
            label_zoom4.place(x=550, y=250)
            # Area for zoom5
            # label_zoom5= Label(root)
            # label_zoom5.place(x=1150,y=562)

            _, frame2 = cap2.read()
            # frame2 = cv2.flip(frame2, 1)
            frame2 = frame2[offset_Ly:frame2.shape[0] -
                            offset_Ry, 0:frame2.shape[1]-offset_Lx]

            img_stack_HD = np.hstack([frame2])
            IMGheight, IMGwidth, _ = img_stack_HD.shape
            ratio = (IMGheight - 350) / IMGheight
# provisoire !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            resized_image = cv2.resize(img_stack_HD, (int(
                IMGwidth * ratio), IMGheight - 350), interpolation=cv2.INTER_CUBIC)
            cv2image3 = cv2.cvtColor(resized_image, cv2.COLOR_BGR2RGB)
            img2 = Image.fromarray(cv2image3)
            imgtk2 = ImageTk.PhotoImage(image=img2)
            label_workspace.imgtk = imgtk2
            label_workspace.configure(image=imgtk2)

            # imgtk_zoomed1 = zoom(img2,100,100)
            imgtk_zoomed2 = zoom(img2, 340, 250)
            imgtk_zoomed3 = zoom(img2, 250, 250)
            imgtk_zoomed4 = zoom(img2, 150, 300)
            # imgtk_zoomed5 = zoom(img2,300,220)

            # Display the zoomed image1 in the label_zoom1
            # label_zoom1.imgtk = imgtk_zoomed1
            # label_zoom1.configure(image=imgtk_zoomed1)
            # Display the zoomed image2 in the label_zoom2
            label_zoom2.imgtk = imgtk_zoomed2
            label_zoom2.configure(image=imgtk_zoomed2)
            # Display the zoomed image3 in the label_zoom3
            label_zoom3.imgtk = imgtk_zoomed3
            label_zoom3.configure(image=imgtk_zoomed3)
            # Display the zoomed image4 in the label_zoom4
            label_zoom4.imgtk = imgtk_zoomed4
            label_zoom4.configure(image=imgtk_zoomed4)
            # Display the zoomed image5 in the label_zoom5
            # label_zoom5.imgtk = imgtk_zoomed5
            # label_zoom5.configure(image=imgtk_zoomed5)
            zoom_done = 'yes'

        return

    # display the informations about the item
    def Infos():
        lookup = data_frame[(data_frame["Item"] == itm)]
        print(itm, 'infooooooos')
        Longueur = lookup.iloc[0]['Longueur']
        Largeur = lookup.iloc[0]['Largeur']
        Nb_parties = lookup.iloc[0]['Nb parties']
        Potting = lookup.iloc[0]['Potting']
        Cx = lookup.iloc[0]['Cx']
        Cy = lookup.iloc[0]['Cy']
        R = lookup.iloc[0]['R']
        messagebox.showinfo("Informations", "Longueur: " + str(Longueur) +
                            "\nLargeur : " + str(Largeur) +
                            "\nNb Parties: " + str(Nb_parties) +
                            "\nPotting : " + str(Potting) +
                            "\nCx : " + str(Cx) +
                            "\nCy : " + str(Cy) +
                            "\nR : " + str(R))

    # function to close the root window and open the info window in case you want to change the item id or the batch id, it is called in the retour button
    def Retour():
        global zoom_done
        cap0.release()
        cap1.release()
        cap2.release()
        root.destroy()
        info.deiconify()

        zoom_done = 'no'

    # for the moment it is not used

    def Swap_Cameras():
        global left_cam
        global right_cam

        if left_cam == 0:
            left_cam = 1
            right_cam = 0
            return
        if left_cam == 1:
            left_cam = 0
            right_cam = 1
            return

    # for the momen it is not used
    def Reset():
        global nb_camera_extra
        global nb_camera_stack

        entry1.delete(0, END)
        entry2.delete(0, END)
        entry3.delete(0, END)
        entry4.delete(0, END)
        nb_camera_extra = 1
        nb_camera_stack = 2

    def Offset1_left_cam():
        global offset_Lx
        offset_Lx = offset_Lx+2

    def Offset2_left_cam():
        global offset_Lx
        if offset_Lx > 0:
            offset_Lx = offset_Lx-2

    def Offset1_right_cam():
        global offset_Rx
        offset_Rx = offset_Rx+2

    def Offset2_right_cam():
        global offset_Rx
        if offset_Rx > 0:
            offset_Rx = offset_Rx-2

    def Offset3_left_cam():
        global offset_Ly
        offset_Ly = offset_Ly+2

    def Offset4_left_cam():
        global offset_Ry
        if offset_Ry > 0:
            offset_Ry = offset_Ry-2

    def Offset3_right_cam():
        global offset_Ry
        offset_Ry = offset_Ry+2

    def Offset4_right_cam():
        global offset_Ry
        if offset_Ry > 0:
            offset_Ry = offset_Ry-2

    # Button "Take Picture"
    label2 = Label(root)
    label2.frame_num = 0
    label2.place(x=400, y=500)
    B = Button(label2, text="Save Image", font="Abadi 10 bold",
               bg="#e85f81", width=10, height=1, command=take_pic)
    B.pack(side="right", fill="x", expand=1)

    # Button "Analyse"
    label4 = Label(root)
    label4.frame_num = 0
    label4.place(x=400, y=550)
    B = Button(label4, text="Analyser", font="Abadi 10",
               bg="#add0f0", width=10, height=1, command=Analyze)
    B.pack(side="right", fill="x", expand=1)

    # Button "Infos"
    label7b = Label(root)
    label7b.frame_num = 0
    label7b.place(x=400, y=600)
    B = Button(label7b, text="plus d\'infos", font="Abadi 10",
               bg="#add0f0", width=10, height=1, command=Infos)
    B.pack(fill="x", expand=1)

    # Button "Retour"
    retour_label = Label(root)
    retour_label.frame_num = 0
    retour_label.place(x=400, y=650)
    retour_button = Button(retour_label, text="Retour", font="Abadi 10",
                           bg="#add0f0", width=10, height=1, command=Retour)
    retour_button.pack(fill="y", expand=1)

    # Button "Swap Camera"
    label8 = Label(root)
    label8.frame_num = 0
    # label8.grid(row=16, column=3, sticky=W, padx=15, pady=5)
    label8.place(x=350, y=268)

    B = Button(label8, text="Swap cameras", font="Abadi 10",
               bg="#add0f0", width=15, height=1, command=Swap_Cameras)
    B.pack(fill="x", expand=1)

    # #Buttons Offset Left Camera
    f10 = Frame(root)
    f10.grid(row=6, column=5, rowspan=4, sticky=W, padx=20, pady=5)

    b15 = Label(f10, text="CAM 0 Set Up", font="Abadi 8")
    # b15.grid(row=1, column=1, columnspan=3, padx=0, pady=3)

    b11 = Button(f10, text="\u2190", font="Abadi 10", bg="#add0f0",
                 width=2, height=1, command=Offset1_left_cam)
    b12 = Button(f10, text="\u2192", font="Abadi 10", bg="#add0f0",
                 width=2, height=1, command=Offset2_left_cam)
    b13 = Button(f10, text="\u2191", font="Abadi 10", bg="#add0f0",
                 width=2, height=1, command=Offset3_left_cam)
    b14 = Button(f10, text="\u2193", font="Abadi 10", bg="#add0f0",
                 width=2, height=1, command=Offset4_left_cam)

    # b11.grid(sticky=W, row=3, column=1)
    # b12.grid(sticky=W, row=3, column=3, padx=5, pady=5)
    # b13.grid(sticky=W, row=2, column=2)
    # b14.grid(sticky=W, row=5, column=2)

    # #Buttons Offset Right Camera
    b25 = Label(f10, text="CAM 1 Set Up", font="Abadi 8")
    # b25.grid(row=1, column=4, columnspan=3, padx=0, pady=3)

    b21 = Button(f10, text="\u2190", font="Abadi 10", bg="#add0f0",
                 width=2, height=1, command=Offset1_right_cam)
    b22 = Button(f10, text="\u2192", font="Abadi 10", bg="#add0f0",
                 width=2, height=1, command=Offset2_right_cam)
    b23 = Button(f10, text="\u2191", font="Abadi 10", bg="#add0f0",
                 width=2, height=1, command=Offset3_right_cam)
    b24 = Button(f10, text="\u2193", font="Abadi 10", bg="#add0f0",
                 width=2, height=1, command=Offset4_right_cam)

    # b21.grid(sticky=W, row=3, column=4, padx=5, pady=5)
    # b22.grid(sticky=W, row=3, column=6)
    # b23.grid(sticky=W, row=2, column=5)
    # b24.grid(sticky=W, row=5, column=5)

    # Title input 1 - Item Number
    label6a = Label(root, text="Item Number", font="Abadi 10")
    label6a.grid(sticky=W, row=12, column=0, padx=15, pady=5)

    # Title input 2 - Batch Number
    label6a = Label(root, text="Batch Number", font="Abadi 10")
    label6a.grid(sticky=W, row=13, column=0, padx=15, pady=5)

    # Title input 3 - Config
    label6a = Label(root, text="Config", font="Abadi 10")
    label6a.grid(sticky=W, row=14, column=0, padx=15, pady=5)

    # Title input 4 - Potting
    label6a = Label(root, text="Potting", font="Abadi 10")
    label6a.grid(sticky=W, row=15, column=0, padx=15, pady=5)

    # #Title input 5 - Position
    # label6a = Label(root, text = "Position", font="Abadi 10")
    # label6a.grid(sticky = W, row=16, column=1, padx = 5, pady = 100)

    # #Title input 5 - User
    # label6f = Label(root, text = "User", font="Abadi 10")
    # label6f.grid(sticky = W, row=16, column=1, padx = 10, pady = 5)

    # label6g = Label(root, text = os.getlogin(), font="Abadi 10")
    # label6g.grid(sticky = W, row=16, column=2, padx = 10, pady = 5)

    # Search program based on item number

    # Input Boxes
    if autres == "no":
        lookup = data_frame[(data_frame["Item"] == itm)]
    elif autres == "yes":
        lookup = data_frame2[(data_frame2["Item"] == itm)]

    camera_activation(lookup.iloc[0]['cam1'])

    entry1 = Entry(root, width=18, font="arial 12")
    entry1.insert(0, lookup.iloc[0]['Item'])
    entry1.configure(state='disabled')
    entry1.config(width='15')

    entry2 = Entry(root, width=18, font="arial 12")
    entry2.insert(0, batch)
    entry2.configure(state='disabled')
    entry2.config(width='15')

    entry3 = Entry(root, width=18, font="arial 12")
    entry3.insert(0, lookup.iloc[0]['Config'])
    global config
    config = lookup.iloc[0]['Config']
    entry3.configure(state='disabled')
    entry3.config(width='15')

    entry4 = Entry(root, width=18, font="arial 12")
    entry4.insert(0, lookup.iloc[0]['Potting'])
    entry4.configure(state='disabled')
    entry4.config(width='15')

    entry1.grid(sticky=W, row=12, column=1, padx=0, pady=5)
    entry2.grid(sticky=W, row=13, column=1, padx=0, pady=5)
    entry3.grid(sticky=W, row=14, column=1, padx=0, pady=5)
    entry4.grid(sticky=W, row=15, column=1, padx=0, pady=5)

    show_frames()
    root.bind("p", key_pressed)
    root.mainloop()


def save_itm(item, batch, conf):
    # Create a DataFrame with the data you want to insert
    if conf == 1:
        cam1 = "yes"
    else:
        cam1 = "no"
    data = {'Operateur': [os.getlogin()],
            'Item': [item],
            'Customer': [" "],
            'Batch': [batch],
            'Config': [conf],
            'Program': [" "],
            'Potting': [" "],
            'Cam1': [cam1]
            }
    df2 = pd.DataFrame(data)

    # Load the existing Excel file
    excel_file = 'C:/Users/louay/Desktop/autres.xlsx'
    sheet_name = 'AUTRES'
    existing_data = pd.read_excel(excel_file, sheet_name=sheet_name)

    # Reset the index of the existing_data DataFrame
    existing_data.reset_index(drop=True, inplace=True)

    # Append the new data to the existing data
    new_data = pd.concat([existing_data, df2], ignore_index=True)

    # Save the updated data back to the Excel file
    new_data.to_excel(excel_file, sheet_name=sheet_name, index=False)

# function to fill the format for the testing process


def fill_formtest():
    ItemID.delete(0, END)
    Batch.delete(0, END)
    ItemID.insert(0, 'M500100')
    Batch.insert(0, '1234568911234')
    check_item()

# conditions checking for input


def check_item():
    global batch
    global customer
    item = ItemID.get()
    batch = Batch.get()
    customer = Customer.get()

    if item == "":
        messagebox.showwarning("Champ non renseigné",
                               "Merci de remplir l'Item Number")
        return

    elif not item in data_frame['Item'].values:
        msg_box = messagebox.askyesno(
            "Item pas enregistré", "Vous voulez Continuer avec cette ID ?")
        if msg_box:
            msg_box2 = messagebox.askquestion(
                "Configuration de la caméra", "est-ce que la taille de la piece dépasse 40cm ?")
            if msg_box2 == "yes":
                print('conf1')
                global itm
                conf = 1
                store_item(item)
                itm = store_item(item)
                save_itm(item, batch, conf)
                open_root_window()
            if msg_box2 == "no":
                print('conf2')
                conf = 2
                store_item(item)
                itm = store_item(item)
                save_itm(item, batch, conf)
                open_root_window()
        else:

            return

    elif batch == "":
        messagebox.showwarning("Champ non renseigné",
                               "Merci de remplir le Batch Number")
        return

    elif item in data_frame['Item'].values:
        if len(item) != 7 or len(batch) != 13:
            # global itm
            msg_box = messagebox.askyesno(
                "Format", "Format de l'item ou du batch non reconnu.\n Voulez-vous continuer.")
            if msg_box:
                store_item(item)
                itm = store_item(item)
                open_root_window()
            else:
                return
        else:
            store_item(item)
            itm = store_item(item)
            open_root_window()

# stores the item number in a global variable to be used later in the root window

# def save_itm(item):


def store_item(item):
    global itm
    itm = item
    return itm


# callback function for the item number that will display the info on the info window
def callback_itemId(id_sv):
    ItemID.get()
    lookup = data_frame[(data_frame["Item"] == id_sv.get())]
    if lookup.shape[0] == 1:
        Customer.delete(0, END)  # deletes the current value
        Customer.insert(0, lookup.iloc[0]['Customer'])
        Customer.configure(state='disabled')

        program_label2.delete(0, END)  # deletes the current value
        program_label2.insert(0, lookup.iloc[0]['Program'])
        program_label2.configure(state='disabled')
        print('cam1 value', lookup['cam1'].values[0])
    else:
        Customer.configure(state='normal')
        Customer.delete(0, END)  # deletes the current value
        program_label2.configure(state='normal')
        program_label2.delete(0, END)  # deletes the current value


id_sv = StringVar()
id_sv.trace("w", lambda name, index, mode, id_sv=id_sv: callback_itemId(id_sv))

# add background
image1 = Image.open("images/background.jpg")
test1 = ImageTk.PhotoImage(image1)
label1 = tkinter.Label(info, image=test1)
label1.image = test1
label1.place(x=0, y=0, relwidth=1, relheight=1)

# title
title_label = Label(info, text="Hexcel Vision Control", font="Abadi 11 bold")
title_label.place(x=120, y=20)  # Adjust the position as needed

# Loading the logo image
logo_hexcel = Image.open("images/logo hexcel 2.png")
logo_resized = logo_hexcel.resize((71, 25))
logo = ImageTk.PhotoImage(logo_resized)
label2 = tkinter.Label(info, image=logo)
label2.image = logo
label2.place(x=40, y=18)

# ITEMID
item_label = Label(info, text='Entrez l\'Item ID')
item_label.place(x=20, y=79)

ItemID = Entry(info, textvariable=id_sv)
info.after(0, ItemID.focus_set)
ItemID.place(x=127, y=80)

# BATCH
batch = None
batch_label = Label(info, text='Entrez le batch')
batch_label.place(x=20, y=119)

Batch = Entry(info)
Batch.place(x=127, y=120)


# Title input 3 - Customer
customer_label = Label(info, text="Customer", font="Abadi 10")
customer_label.place(x=20, y=159)

# output - Customer
customer = None
Customer = Entry(info, width=18, font="Abadi 10")
Customer.place(x=127, y=159)
Customer.config(state='disabled')

# Title input 4 - Program
program_label = Label(info, text="Program", font="Abadi 10")
program_label.place(x=20, y=199)

# output - Program
program_label2 = Entry(info, width=18, font="Abadi 10")
program_label2.place(x=127, y=199)
program_label2.config(state='disabled')

# Title input 5 - User
user_label = Label(info, text="Operateur", font="Abadi 10")
user_label.place(x=20, y=239)

login_label = Label(info, text=os.getlogin(), font="Abadi 10")
login_label.place(x=127, y=239)

# Button to open the root window as a popup
open_button = Button(info, text="Ouvrir l'application",
                     command=check_item, font="Abadi 10 bold", fg='#0047AB')
open_button.place(x=120, y=280)

#################################################
# once you finish testing comment this button
# Button to test
TEST = Button(info, text="TEST",
              command=fill_formtest, font="Abadi 10 bold", fg='#0047AB')
TEST.place(x=160, y=310)
#################################################


info.resizable(width=False, height=False)
info.mainloop()
