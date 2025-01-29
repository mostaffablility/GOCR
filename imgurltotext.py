import os
import time
from tkinter import Tk
from tkinter.filedialog import askdirectory
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from html2docx import html2docx


# Function to upload an image URL to Google Lens
def upload_image_to_google_lens(driver, image_url):
    try:
        # Wait for the input field
        input_field = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[text='text']"))
        )
        print(f"Found input field: {input_field}")

        # Clear and send the image URL
        input_field.clear()
        input_field.send_keys(image_url)
        input_field.send_keys(Keys.RETURN)
        print(f"Uploaded image URL: {image_url}")
    except Exception as e:
        print(f"Error uploading image to Google Lens: {e}")


# Function to save RTMDre element as an HTML file with visible aria-label text
def save_element_as_html(driver, css_selector, output_html_file):
    try:
        # Locate the RTMDre container
        element = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, css_selector))
        )
        print("RTMDre container found.")

        # Wait until the RTMDre container has child elements (fully loaded)
        WebDriverWait(driver, 20).until(lambda d: len(element.find_elements(By.CSS_SELECTOR, '.lv6PAb')) > 5)
        print("RTMDre container fully loaded with child elements.")

        # Get the dimensions of the container (width and height)
        container_dimensions = driver.execute_script(
            """
            const container = arguments[0];
            const rect = container.getBoundingClientRect();
            return {
                width: rect.width,
                height: rect.height
            };
            """,
            element
        )
        width = container_dimensions['width']
        height = container_dimensions['height']
        print(f"Container dimensions: Width={width}px, Height={height}px")

        # Execute JavaScript to keep only specified elements within the RTMDre container
        modified_html = driver.execute_script(
            """
            const container = arguments[0];
            const elementsToKeep = Array.from(container.querySelectorAll('.lv6PAb.PyT1Q'));
            const allChildren = Array.from(container.children);
            allChildren.forEach(el => {
                if (!elementsToKeep.includes(el)) el.remove(); // Keep only specified elements
            });

            // Inject visible text for aria-label
            elementsToKeep.forEach(el => {
                const span = document.createElement('span');
                span.textContent = el.getAttribute('aria-label');
                span.style.cssText = `
                    color: black;
                    font-size: 14px;
                    position: absolute;
                    left: 0;
                    top: 0;
                    z-index: 100;
                    background-color: rgba(255, 255, 255, 0.8);
                `;
                el.appendChild(span);
            });

            return container.outerHTML;
            """,
            element
        )
        print("JavaScript integration completed for RTMDre container.")

        # Wrap the modified RTMDre container in a new outer container with the same dimensions as the original
        outer_container_html = f"""
        <div style="
            width: {width}px;
            height: {height}px;
            position: relative;
            overflow: hidden;
            background-color: white;
            border: 1px solid #ddd;">
            {modified_html}
        </div>
        """

        # Clone the document's styles
        styles = driver.execute_script("return document.head.outerHTML;")

        # Check for the text direction (RTL or LTR)
        direction = driver.execute_script("return getComputedStyle(document.body).direction;")
        direction_style = "direction: rtl; text-align: right;" if direction == "rtl" else "direction: ltr; text-align: left;"

        # Create the full HTML content
        full_html = f"""
        <!DOCTYPE html>
        <html lang="en">
        <head>{styles}</head>
        <body style="{direction_style}; margin: 0; padding: 0;">
            {outer_container_html}
        </body>
        </html>
        """

        # Save the HTML to a file
        with open(output_html_file, 'w', encoding='utf-8') as file:
            file.write(full_html)

        print(f"HTML content saved to {output_html_file}")

    except Exception as e:
        print(f"Error saving element as HTML: {e}")



# Function to convert HTML to DOCX
def convert_html_to_docx(input_html_file, output_docx_file):
    try:
        with open(input_html_file, 'r', encoding='utf-8') as file:
            html_content = file.read()

        # Convert HTML to DOCX
        docx_content = html2docx(html_content)
        with open(output_docx_file, 'wb') as file:
            file.write(docx_content)

        print(f"DOCX file saved to {output_docx_file}")
    except Exception as e:
        print(f"Error converting HTML to DOCX: {e}")


# Function to process images on the web server
def process_images_with_google_lens(base_url, output_folder):
    driver = webdriver.Chrome()

    # Example: List of image files on your web server
    image_files = [
        "image1.jpg",
        "image2.jpg",
        "image3.jpg"
    ]

    for file_name in image_files:
        image_url = f"{base_url}/{file_name}"
        print(f"Processing: {file_name} (URL: {image_url})")

        try:
            # Navigate to Google Lens
            driver.get("https://www.google.com/?olud")
            time.sleep(2)

            # Upload the image to Google Lens
            upload_image_to_google_lens(driver, image_url)

            # Save the RTMDre element as an HTML file
            html_file = os.path.join(output_folder, f"{os.path.splitext(file_name)[0]}.html")
            save_element_as_html(driver, ".RTMDre", html_file)

            # Convert the HTML file to DOCX
            docx_file = os.path.join(output_folder, f"{os.path.splitext(file_name)[0]}.docx")
            convert_html_to_docx(html_file, docx_file)

            print(f"Processed {file_name}, saved as {docx_file}")

        except Exception as e:
            print(f"Error processing {file_name}: {e}")

    driver.quit()


# Main function
if __name__ == '__main__':
    # Base URL for images on your web server
    base_url = "http://mosishoki18.unaux.com/img"

    # Prompt the user to select a folder for saving the output
    root = Tk()
    root.withdraw()
    print("Select a folder to save the output files...")
    output_folder = askdirectory(title="Select Output Folder")

    if not output_folder:
        print("No folder selected. Exiting.")
        exit()

    print(f"Output folder selected: {output_folder}")
    os.makedirs(output_folder, exist_ok=True)

    print(f"Processing images from {base_url}")
    process_images_with_google_lens(base_url, output_folder)
    print(f"Processing completed. Output saved to '{output_folder}'.")
