import comtypes.client
import os

def ppt_to_pdf(input_file_path, output_file_path, format_type=32):
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1

    if not os.path.isabs(input_file_path):
        input_file_path = os.path.abspath(input_file_path)
    
    if output_file_path[-4:].lower() != '.pdf':
        output_file_path += ".pdf"
    
    deck = powerpoint.Presentations.Open(input_file_path)
    deck.SaveAs(output_file_path, format_type)
    deck.Close()
    powerpoint.Quit()

input_file_path = r'C:\Users\m.x\Desktop\ppt\123.ppt'
output_file_path = r'C:\Users\m.x\Desktop\pdf\123.pdf'
ppt_to_pdf(input_file_path, output_file_path)
