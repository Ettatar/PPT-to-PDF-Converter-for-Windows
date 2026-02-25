import comtypes.client
import os

def batch_convert_ppt_to_pdf():
    # Get the current directory where the script is located
    current_dir = os.getcwd()
    
    # Initialize the PowerPoint application once
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    
    # Iterate through all files in the current directory
    for file in os.listdir(current_dir):
        # Select only files with .ppt and .pptx extensions
        if file.lower().endswith((".ppt", ".pptx")):
            input_path = os.path.join(current_dir, file)
            
            # Set the output filename (replace extension with .pdf)
            output_filename = os.path.splitext(file)[0] + ".pdf"
            output_path = os.path.join(current_dir, output_filename)
            
            print(f"Converting: {file}...")
            
            try:
                # Open the presentation (WithWindow=False runs it in the background)
                presentation = powerpoint.Presentations.Open(input_path, WithWindow=False)
                
                # 32 is the constant for PDF format in PowerPoint
                presentation.SaveAs(output_path, 32)
                
                presentation.Close()
                print(f"Completed: {output_filename}")
            except Exception as e:
                print(f"Error occurred during conversion of ({file}): {e}")

    # Close the PowerPoint application
    powerpoint.Quit()
    print("\nProcess finished! All presentations have been converted to PDF.")

if __name__ == "__main__":
    batch_convert_ppt_to_pdf()