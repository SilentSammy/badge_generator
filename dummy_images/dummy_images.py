import cv2
import numpy as np
import os

def create_dummy_images(output_dir=".", count=30, size=512):
    """
    Create dummy images with red outline and centered number text.
    
    Args:
        output_dir (str): Directory to save images
        count (int): Number of images to create
        size (int): Image dimensions (size x size)
    """
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    for i in range(1, count + 1):
        # Create a white background image
        img = np.ones((size, size, 3), dtype=np.uint8) * 255
        
        # Add red outline around the edges (5 pixel border)
        border_thickness = 5
        cv2.rectangle(img, 
                     (0, 0), 
                     (size - 1, size - 1), 
                     (0, 0, 255),  # Red color in BGR format
                     border_thickness)
        
        # Add the number in the center
        text = str(i)
        
        # Calculate font size based on image size and text length
        font = cv2.FONT_HERSHEY_SIMPLEX
        font_scale = max(1, size // 100)  # Dynamic font size
        thickness = max(2, size // 150)   # Dynamic thickness
        
        # Get text size to center it
        (text_width, text_height), baseline = cv2.getTextSize(text, font, font_scale, thickness)
        
        # Calculate position to center the text
        x = (size - text_width) // 2
        y = (size + text_height) // 2
        
        # Add text to image (black color)
        cv2.putText(img, text, (x, y), font, font_scale, (0, 0, 0), thickness, cv2.LINE_AA)
        
        # Save the image
        filename = f"dummy_{i:02d}.jpg"
        filepath = os.path.join(output_dir, filename)
        cv2.imwrite(filepath, img)
        
        print(f"Created: {filename}")
    
    print(f"\nSuccessfully created {count} dummy images in '{output_dir}'")

if __name__ == "__main__":
    # Create 30 dummy images in the current directory
    create_dummy_images(output_dir=".", count=30, size=512)