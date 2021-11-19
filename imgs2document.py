#The application is used to create a pdf or docx file from a folder of images, 
#images are appended to the pdf or word file in the order of the filenames

import easygui
import os
import PIL
import fstring
import docx
from fpdf import FPDF
from tqdm import tqdm

def insertImages2WordDoc(img_path=None, doc_path=None):        
        # Get the path to the folder
        if img_path is None:
            image_folder = easygui.diropenbox()
        else:
            image_folder = img_path
        
        # generate the name of the pdf file based on path
        if doc_path is None:
            doc_name = os.path.basename(image_folder) + ".docx"
        else:
            doc_name = doc_path
        
        #check if there is a pdf file with the same name in the folder, delete it if there is
        if os.path.isfile(doc_name):
            os.remove(doc_name)
            
        image_list=[]
        
        #walk through the folder and get all the images    
        for root, dirs, files in os.walk(image_folder):
            for file in files:
                #first conver filaname to lowercase
                file_name = file.lower()
                if file_name.endswith((".jpg", ".png", ".jpeg")):
                    #test if the file is really a picture by reading the first few bytes, if this is not a picture, skip it
                    try:
                        PIL.Image.open(os.path.join(root, file))
                        image_list.append(os.path.join(root, file))
                    except Exception as e:
                        print(f"[Error] : inserting {file} failed, skip! Failure reason: \n        Exception : {e}\n")
                
        #sort the list of images
        image_list.sort()
        print(f"{len(image_list)} images found in the folder {image_folder}:")        
        
        #now we have a list of images, we can create the pdf file
        doc = docx.Document()                        
        for image in tqdm(image_list, desc="Inserting images to document",total=len(image_list)):
           #insert the image to doc, make sure the size of the image is not bigger than the page and keep ratio           
          #in case of any exception, skip the image
            try:                
                p=doc.add_paragraph()
                p.alignment=docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
                r=p.add_run()
                r.add_picture(image, width=docx.shared.Inches(6))
                doc.add_page_break()                                            
            except:
                print(f"{image} is not a picture")
        doc.save(doc_name)        

        print(f"Word file {doc_name} created, {len(image_list)} images added!")
        

def insertImages2PDF(img_path=None, pdf_path=None):
    
    # Get the path to the folder
    if img_path is None:
        image_folder = easygui.diropenbox()
    else:
        image_folder = img_path
    
    # generate the name of the pdf file based on path
    if pdf_path is None:
        pdf_name = os.path.basename(image_folder) + ".pdf"
    else:
        pdf_name = pdf_path
    
    #check if there is a pdf file with the same name in the folder, delete it if there is
    if os.path.isfile(pdf_name):
        os.remove(pdf_name)
        
    image_list=[]    
    #walk through the folder and get all the images    
    for root, dirs, files in os.walk(image_folder):
        for file in files:
            #first conver filaname to lowercase
            file_name = file.lower()            
            if file_name.endswith((".jpg", ".png", ".jpeg")):
                image_list.append(os.path.join(root, file))
            
    #sort the list of images
    image_list.sort()
    print(f"{len(image_list)} images found in the folder {image_folder}:")
    
    layoutOreintation = "L"

    pdf = FPDF('L', 'mm', 'A4')
    
    #define A4 page width and height
    if layoutOreintation == "P":
        A4_page_width = 210
        A4_page_height = 297
    else:
        A4_page_width = 297
        A4_page_height = 210

    pdf.set_auto_page_break(False)        
    img=None
    
    for id, image in tqdm(enumerate(image_list), desc="Inserting images to pdf",total=len(image_list)):
        try:
                        
            #get the width and height of the image by PIL
            
            #use PIL to open the image file based on file extension

            img = PIL.Image.open(image).convert('RGB')
            
            #get dpi information from the image
            
            img = ResizeImage(img,size=8, convert=True)    
            
            #get the width and height of the temp file image by PIL
            
            width_pixel = img.size[0]
            height_pixel = img.size[1]
            dpi = img.info.get('dpi', (300, 300))   

            #convert width and height to mm
            width_mm = int(width_pixel / dpi[0] * 25.4)            
            height_mm = int(height_pixel / dpi[1] * 25.4)            
                                            
            #save img to temp file
            temp_file = f"temp_{id}.jpg"
            img.save(temp_file, "JPEG",dpi=dpi)

            #debug print filename, width and height and dpi
            print(f"{image}  w x h = {width_mm}x{height_mm} dpi={dpi}")
            
            pdf.add_page()            
            
            #calculate the x and y poistion of the image based on page size, make sure the image is in the center of the page
            pos_x = (A4_page_width- width_mm) / 2
            pos_y = (A4_page_height- height_mm) / 2
            
            #insert the image to pdf, make sure the size of the image is not bigger than the page and keep ratio            
            pdf.image(temp_file, pos_x, pos_y, w=width_mm, h=height_mm)
            
            #remove temp file
            os.remove(temp_file)
            img.close()
            img=None
            
            print(f"{image} added to pdf")
              
        except Exception as e:
            print(f"[Exception]: when handling file {image}, Exception happened:\n{e}\n")   
            #remove temp file if exists
            if os.path.isfile(temp_file):
                os.remove(temp_file)
            
            #close img if it's open
            if img is not None:
                img.close()
            continue
        
    pdf.output(pdf_name, "F")
    
    print(f"PDF file {pdf_name} created, {len(image_list)} images added!")

    

def ResizeImage(image,size=6, convert=True):
    '''image is an image which is already opened by PIL.
       size is the size of the image resolution,
       convert is the option if the pic should be converted to RGB if it's RGBA
    '''
    new_height=0
    new_width=0    
    
    #first get width and height of the image
    #open the image
    
    width_pixel = image.size[0]
    height_pixel = image.size[1]
    dpi = image.info.get('dpi', (300, 300))   
      
    #calculate the new width and height
    if width_pixel > height_pixel:
        new_width = int(size * dpi[0])
        new_height = int(new_width * height_pixel / width_pixel)
    else:
        new_height = int(size * dpi[1])
        new_width = int(new_height * width_pixel / height_pixel)
    
    #print the resize info
    print(f"\nResize image from {width_pixel}x{height_pixel} to {new_width}x{new_height}, dpi={dpi}")
    
    #resize the image
    img = image.resize((int(new_width), int(new_height)), PIL.Image.ANTIALIAS)
    
    #convert the image to RGB if it's RGBA
    if convert:
        img = img.convert('RGB')
    
    return img

if __name__ == "__main__":
    #First ask the user if he wants to create a pdf or a word file by using easygui choicebox
    choice = easygui.choicebox("Do you want to create a pdf or a word file from images?", choices=["PDF", "Word"])

    if choice == "PDF":        
        insertImages2PDF()
    elif choice == "Word":
        insertImages2WordDoc()