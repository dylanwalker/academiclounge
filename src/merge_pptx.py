import os
import win32com.client

# Define paths
path_to_presentations = r"C:\path\to\presentations"  # Replace with your folder path
path_to_presentations = os.path.join(os.path.dirname(__file__), os.pardir, "slides")


path_to_output = r"C:\path\to\output"  # Replace with your output folder path
path_to_output = os.path.join(os.path.dirname(__file__), os.pardir, "slideshow")


if not os.path.exists(path_to_output):
    os.makedirs(path_to_output)

# Initialize PowerPoint application
powerpoint = win32com.client.Dispatch("PowerPoint.Application")
powerpoint.Visible = 1  # Make PowerPoint visible (optional)

# Create a new presentation
merged_presentation = powerpoint.Presentations.Add()

# Get all .pptx files in the folder
pptx_files = [os.path.join(path_to_presentations, f) for f in os.listdir(path_to_presentations) if f.endswith('.pptx')]

for pptx_file in pptx_files:
    # Open the current presentation
    curr_presentation = powerpoint.Presentations.Open(pptx_file, WithWindow=False)
    
    # Loop through slides and copy them to the merged presentation
    for slide_index in range(1, curr_presentation.Slides.Count + 1):
        curr_presentation.Slides(slide_index).Copy()
        merged_presentation.Slides.Paste(-1)  # Paste at the end of the merged presentation

    # Close the current presentation
    curr_presentation.Close()
    print(f"Finished processing {pptx_file}")

# Save the merged presentation
output_file = os.path.join(path_to_output, "merged_presentation.pptx")
merged_presentation.SaveAs(output_file)

# Close PowerPoint
merged_presentation.Close()
powerpoint.Quit()

print(f"Saved merged presentation to {output_file}")