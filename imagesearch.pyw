from deepface import DeepFace
import os
import shutil
import cv2 
import numpy as np
from deepface.detectors import FaceDetector
import tensorflow as tf
import traceback
import tkinter as tk
tf_version = tf.__version__
tf_major_version = int(tf_version.split(".")[0])
tf_minor_version = int(tf_version.split(".")[1])
if tf_major_version == 1:
	import keras
	from keras.preprocessing.image import load_img, save_img, img_to_array
	from keras.applications.imagenet_utils import preprocess_input
	from keras.preprocessing import image
elif tf_major_version == 2:
	from tensorflow import keras
	from tensorflow.keras.preprocessing.image import load_img, save_img, img_to_array
	from tensorflow.keras.applications.imagenet_utils import preprocess_input
	from tensorflow.keras.preprocessing import image


def detect_face(result,img):

    img_region = [0, 0, img.shape[1], img.shape[0]]
  
    try:
        detected_face, img_region = result
    except: #if detected face shape is (0, 0) and alignment cannot be performed, this block will be run
        detected_face = None

    if (isinstance(detected_face, np.ndarray)):
        return detected_face, img_region
    else:
        if detected_face == None:
            enforce_detection = False
            if enforce_detection != True:
                return img, img_region
            else:
                raise ValueError("Face could not be detected. Please confirm that the picture is a face photo or consider to set enforce_detection param to False.")



def preprocess_face(img, output, target_size=(224, 224), grayscale = False, enforce_detection = False, detector_backend = 'opencv', return_region = False, align = True):

	#img might be path, base64 or numpy array. Convert it to numpy whatever it is.
	img = DeepFace.functions.load_image(img)
	base_img = img.copy()

	img, region = output

	#--------------------------

	if img.shape[0] == 0 or img.shape[1] == 0:
		if enforce_detection == True:
			raise ValueError("Detected face shape is ", img.shape,". Consider to set enforce_detection argument to False.")
		else: #restore base image
			img = base_img.copy()

	#--------------------------

	#post-processing
	if grayscale == True:
		img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

	#---------------------------------------------------
	#resize image to expected shape

	# img = cv2.resize(img, target_size) #resize causes transformation on base image, adding black pixels to resize will not deform the base image

	if img.shape[0] > 0 and img.shape[1] > 0:
		factor_0 = target_size[0] / img.shape[0]
		factor_1 = target_size[1] / img.shape[1]
		factor = min(factor_0, factor_1)

		dsize = (int(img.shape[1] * factor), int(img.shape[0] * factor))
		img = cv2.resize(img, dsize)

		# Then pad the other side to the target size by adding black pixels
		diff_0 = target_size[0] - img.shape[0]
		diff_1 = target_size[1] - img.shape[1]
		if grayscale == False:
			# Put the base image in the middle of the padded image
			img = np.pad(img, ((diff_0 // 2, diff_0 - diff_0 // 2), (diff_1 // 2, diff_1 - diff_1 // 2), (0, 0)), 'constant')
		else:
			img = np.pad(img, ((diff_0 // 2, diff_0 - diff_0 // 2), (diff_1 // 2, diff_1 - diff_1 // 2)), 'constant')

	#------------------------------------------

	#double check: if target image is not still the same size with target.
	if img.shape[0:2] != target_size:
		img = cv2.resize(img, target_size)

	#---------------------------------------------------

	#normalizing the image pixels

	img_pixels = image.img_to_array(img) #what this line doing? must?
	img_pixels = np.expand_dims(img_pixels, axis = 0)
	img_pixels /= 255 #normalize input in [0, 1]

	#---------------------------------------------------

	if return_region == True:
		return img_pixels, region
	else:
		return img_pixels


def checkresult(result):
    try:
        detected_face, img_region = result
    except: #if detected face shape is (0, 0) and alignment cannot be performed, this block will be run
        detected_face = None
    if detected_face is None:
        return False
    if detected_face.shape[0] == 0 or detected_face.shape[1] == 0:
        return False
    return True




testfolder = r'E:\Data Copy\Image_Search_Program\Data'
inputfolder = r'E:\Data Copy\Image_Search_Program\Input'
outputfolder = r'E:\Data Copy\Image_Search_Program\Output'


def verify(sourceembedding,path,face_detector, detector,file):
    embeddings = []
    img = cv2.imread(path)
    results = FaceDetector.detect_faces(face_detector, detector, img, True)
    input_shape_x, input_shape_y = DeepFace.functions.find_input_shape(model)
    ids = []
    resultok = []
    for result in results:
        resultok.append(checkresult(result))
    if any(resultok) == False:
        print('using other detector')
        if detector == 'retinaface':
            detector = 'mtcnn'
        else:
            detector = 'retinaface'
        face_detector = FaceDetector.build_model(detector)
        results = FaceDetector.detect_faces(face_detector, detector, img, True)

    for result in results:
        output = detect_face(result,img)
        output2 = preprocess_face(img,output, target_size=(input_shape_y, input_shape_x))
        output3 = DeepFace.functions.normalize_input(img = output2, normalization = 'base')
        if "keras" in str(type(model)):
            #new tf versions show progress bar and it is annoying
            targetembedding = model.predict(output3, verbose=0)[0].tolist()
        else:
            #SFace is not a keras model and it has no verbose argument
            targetembedding = model.predict(output3)[0].tolist()
        distance = DeepFace.dst.findCosineDistance(sourceembedding, targetembedding)
        distance = np.float64(distance)
        threshold = DeepFace.dst.findThreshold(model_name, 'cosine')
        if distance <= threshold:
            identified = True
            foundembeddings.append(targetembedding)
        else:
            identified = False
        print({"identified":identified})
        ids.append(identified)
        embeddings.append({'path':path,'file':file,'embedding':targetembedding})
    return {"verified":any(ids),'embeddings':embeddings}

def checkfolder():

    for file in os.listdir(testfolder):
        test = testfolder + "\\" + file
        try:
            check = cv2.imread(test)
        except:
            continue
        if check is None:
            continue
        
        
        for sourceembedding in sourceembeddings:
            result = verify(sourceembedding,test,face_detector, detector, file)
            testembeddings.extend(result['embeddings'])
            # result = DeepFace.verify(img1_path = img , img2_path = test ,enforce_detection =False, model = builtmodel,  detector_backend = 'opencv')
            # result2 = DeepFace.verify(img1_path = img , img2_path = test ,enforce_detection =False, model = builtmodel,  detector_backend = 'mtcnn')
            if result['verified'] == True : #or result2['verified'] == True 
                outputpath = outputfolder + "\\" + file
                shutil.copyfile(test, outputpath)
                print (result['verified'])

def checkembeddings():
        
    for sourceembedding in sourceembeddings:
        for a in testembeddings:
            path = a['path']
            file = a['file']
            targetembedding = a['embedding']
            distance = DeepFace.dst.findCosineDistance(sourceembedding, targetembedding)
            distance = np.float64(distance)
            threshold = DeepFace.dst.findThreshold(model_name, 'cosine')
            if distance <= threshold:
                outputpath = outputfolder + "\\" + file
                if file not in os.listdir(outputfolder):
                    shutil.copyfile(path, outputpath)
                    foundembeddings.append(targetembedding)

model_name = "Facenet512"
model = DeepFace.build_model(model_name)
sourceembeddings = []
foundembeddings = []
testembeddings = []
detector = 'retinaface'
face_detector = FaceDetector.build_model(detector)


def facecheck():
    try:
        global foundembeddings, sourceembeddings, testembeddings
        for input in os.listdir(inputfolder):
            srcimg = inputfolder + "\\" + input
            try:
                check = cv2.imread(srcimg)
            except:
                continue
            if check is None:
                continue
            try:
                embedding = DeepFace.represent(srcimg, model_name, model = model, enforce_detection = True, detector_backend = 'retinaface')
            except:
                embedding = DeepFace.represent(srcimg, model_name, model = model, enforce_detection = False, detector_backend = 'mtcnn')
                print('used mtcnn')
            sourceembeddings.append(embedding)
        checkfolder()
        
        sourceembeddings.extend(foundembeddings)
        foundembeddings =[]
        print('length of source embeddings', len(sourceembeddings))
        checkembeddings()
        sourceembeddings.extend(foundembeddings)
        foundembeddings =[]
        print('length of source embeddings', len(sourceembeddings))
        checkembeddings()


    except:
        print(srcimg,test)
        print(traceback.format_exc())


if __name__ == "__main__":
    try:
        window = tk.Tk()
        window.title('Face Search tool')
        window.geometry("400x150")
        frame = tk.Frame(window)
        frame.grid(row=0, column=0)
        button3 = tk.Button(frame, text='Search Face', command=facecheck)
        button3.grid(row=0, column=1)
  

        window.mainloop()
    except:
        window.destroy()
