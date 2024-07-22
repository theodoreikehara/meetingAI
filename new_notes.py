import os
import threading
from tkinter import Tk, filedialog, Button, Label
from tkinter.ttk import Progressbar, Combobox
import openai
from docx import Document
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv


# Load the .env file
load_dotenv()

AI_API_KEY = os.getenv('AI_API_KEY')
EMAIL_API_KEY = os.getenv('EMAIL_API_KEY')

# New functions to save and load the last used email
def save_last_used_email(email):
    email = email.strip()  # Remove leading/trailing whitespace
    if not email:  # If the email is empty after stripping, do nothing
        return
    try:
        with open('last_email_temp.txt', 'r+') as file:
            emails = set(file.read().splitlines())
            if email not in emails:
                file.write(f"{email}\n")
    except FileNotFoundError:
        with open('last_email_temp.txt', 'w') as file:
            file.write(f"{email}\n")

def load_last_used_email():
    try:
        with open('last_email_temp.txt', 'r') as file:
            emails = file.read().splitlines()
            # Filter out empty lines and lines with only whitespace
            emails = [email for email in emails if email.strip()]
            return emails if emails else []
    except FileNotFoundError:
        return []
    
def update_email_combobox():
    # Reload the last used emails
    last_used_emails = load_last_used_email()
    # Update the Combobox values
    email_combobox['values'] = last_used_emails
    # Set the default value to the last used email if available
    if last_used_emails:
        email_combobox.set(last_used_emails[-1])
    
# Function to read docx file 
def read_docx(file_path):
    doc = Document(file_path)
    text = "\n".join(para.text for para in doc.paragraphs)
    return text

def read_vtt(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            # You can add additional processing here if needed
            return file.read()
    except Exception as e:
        print(f"Error reading VTT file: {e}")
        return None

def update_progress_bar():
    progress_bar.pack(pady=10)  # Make the progress bar visible
    progress_bar['mode'] = 'indeterminate'
    progress_bar.start(10)
    status_label.config(text="Generating summary, please wait...")

def stop_progress_bar():
    progress_bar.stop()
    progress_bar['mode'] = 'determinate'
    status_label.config(text="Finished generating summary!")
    progress_bar.pack_forget()  # Hide the progress bar

def process_selected_file(file_path, recipient_email):
    file_extension = os.path.splitext(file_path)[1]
    
    if file_extension == ".docx":
        text = read_docx(file_path)
    elif file_extension == ".mp3":
        text = transcribe_audio(file_path)
    elif file_extension == ".vtt":
        text = read_vtt(file_path)
    else:
        root.after_idle(lambda: status_label.config(text="Unsupported file type selected."))
        return
    
    if text is None:  # If transcription, reading or VTT processing failed
        root.after_idle(lambda: status_label.config(text="Failed to process the file. Please try again."))
        return
    
    # Now generate_summary_thread function will be called from here
    generate_summary_thread(text, recipient_email)
    
# Function to handle file selection and processing
def process_file():
    # Allow selection of docx, mp3, and vtt files
    file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx"), ("Audio Files", "*.mp3"), ("VTT Files", "*.vtt")])
    recipient_email = email_combobox.get().strip()  # Ensure whitespace is trimmed
    if not file_path:
        status_label.config(text="Please select a document, audio file, or VTT file to process.")
        return
    if not recipient_email:
        status_label.config(text="Please fill in the recipient's email address.")
        return
    save_last_used_email(recipient_email)  # Save the email when processing
    
    # Start the progress bar as soon as file is selected and before processing
    update_progress_bar()
    
    # Processing is moved to a separate thread to keep UI responsive
    threading.Thread(target=lambda: process_selected_file(file_path, recipient_email)).start()

def transcribe_audio(file_path):
    client = openai.OpenAI(api_key = AI_API_KEY)
    try:
        with open(file_path, "rb") as audio_file:
            transcript = client.audio.transcriptions.create(
                model="whisper-1", 
                file=audio_file, 
                response_format="text"
            )
        return transcript
    except Exception as e:
        print(f"Error during transcription: {e}")
        status_label.config(text=f"Error during transcription: {e}")
        return None

def generate_summary_thread(text, recipient_email):
    try:
        # Existing code to generate summary and send email
        response = generate_summary(text)
        send_email(recipient_email, response)
        print(response)  # or display it in the UI
        stop_progress_bar()
    finally:
        # Safely stop and hide the progress bar on the main thread after processing is complete
        root.after_idle(stop_progress_bar)
        # Update the email dropdown on the main thread
        root.after_idle(update_email_combobox)

# Function to generate the summary (your existing code)
def generate_summary(text):

    client = openai.OpenAI(api_key=AI_API_KEY)

    try:
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are a highly skilled AI trained in language comprehension and summarization. I would like you to read the following text and summarize it into a concise meeting notes. Aim to retain the most important points, providing a coherent and readable notes that could help a person understand the main points of the discussion without needing to read the entire text. Please avoid unnecessary details or tangential points."},
                {"role": "system", "content": "Please generate meeting NOTES, TASKS and KEY DECISIONS from the transcript. \nGenerate in this format:\n- Notes:\n- Tasks:\n- Key Decisions:"},
                {"role": "user", "content": text}
            ]
        )
    except openai.OpenAIError as e:
        print('Long text detected:')
        if 'context_length_exceeded' in str(e):
            # Split the text into two parts
            half_length = len(text) // 2
            first_half = text[:half_length]
            second_half = text[half_length:]

            # Try generating summary for the first half
            first_response = generate_summary(first_half)

            # Try generating summary for the second half
            second_response = generate_summary(second_half)

            # Combine the responses
            combined_response = first_response + '\n' + second_response
            final_response = generate_summary(combined_response)
            
            return final_response
        else:
            print(f"Error during API call: {e}")
            return None
    else:
        return response.choices[0].message.content

def send_email(receiver_address, content):

    # Dynamically creates meeting subject

    client = openai.OpenAI(api_key=AI_API_KEY)

    response = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[
            {"role": "system", "content": "You are an assistant for creating a Subject line for an email given the meeting notes"},
            {"role": "system", "content": "Make the subject line concise and only capture the main and important points"},
            {"role": "system", "content": "Don't use the words subject or subject line"},
            {"role": "user", "content": content}
        ]
    )

    subject = response.choices[0].message.content
    
    # Email credentials
    sender_address = "mscomeetingsummary@gmail.com"
    # sender_pass = os.getenv("EMAIL_API_KEY")  

    sender_pass = EMAIL_API_KEY
    #

    # Setup the MIME
    message = MIMEMultipart()
    message['From'] = sender_address
    message['To'] = receiver_address
    message['Subject'] = subject
    message.attach(MIMEText(content, 'plain'))

    # Create SMTP session for sending the mail
    try:
        session = smtplib.SMTP('smtp-relay.brevo.com', 587)  # Use your SMTP server
        session.starttls()  # Enable security
        session.login(sender_address, sender_pass)  # Login with email and password
        text = message.as_string()
        session.sendmail(sender_address, receiver_address, text)
        session.quit()
        print("Email sent successfully")
    except Exception as e:
        print(f"Error: {e}")
        root.after_idle(lambda: status_label.config(text=f"Error: {e}"))

# Set up the main application window
root = Tk()
root.title("Document Summary Generator Beta (0.3.3)")
root.geometry("400x250")

# Add a label and entry widget for email input
email_label = Label(root, text="Recipient Email:")
email_label.pack(pady=5)
# Load the last used emails and prepare the Combobox
last_used_emails = load_last_used_email()
email_combobox = Combobox(root, width=27)

# Set the Combobox values to the list of last used emails
email_combobox['values'] = last_used_emails

# Set the default value to the last used email if available
if last_used_emails:
    email_combobox.set(last_used_emails[-1])  # Default to the last email in the list

email_combobox.pack(pady=5)

# Add a button to open the file dialog
button = Button(root, text="Open Document", command=process_file)
button.pack(pady=20)

# Add a status label
status_label = Label(root, text="")
status_label.pack(pady=5)

# Add a progress bar
progress_bar = Progressbar(root, orient='horizontal', length=300, mode='determinate')
progress_bar.pack(pady=10)
progress_bar.pack_forget()  # Initially hide the progress bar

# Run the application
root.mainloop()