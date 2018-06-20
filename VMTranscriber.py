#------------------------------------------------------------------------------------------
#Import statements
#------------------------------------------------------------------------------------------

import time
import datetime
import boto3
import json
import configparser
import os
import sys
import threading
import win32com.client
import win32com
from pathlib import Path
from urllib.request import urlopen

#------------------------------------------------------------------------------------------
#Transcribe clients for boto3
#------------------------------------------------------------------------------------------

transcribe = boto3.client('transcribe', aws_access_key_id="AKIAIGL725O7PZ2ZFXDQ", aws_secret_access_key="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", region_name="us-east-1")
s3 = boto3.resource('s3', aws_access_key_id="AKIAIGL725O7PZ2ZFXDQ", aws_secret_access_key="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", region_name="us-east-1")
s3Client = boto3.client('s3', aws_access_key_id="AKIAIGL725O7PZ2ZFXDQ", aws_secret_access_key="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", region_name="us-east-1")

#------------------------------------------------------------------------------------------
#Empty variables to handle input name
#------------------------------------------------------------------------------------------

file_name = ""
transcript_url = ""
transcript_text = ""

#------------------------------------------------------------------------------------------
#User prompt and array initialized
#------------------------------------------------------------------------------------------
'''
input_value = input("Enter the name of your MP3 file you would like transcribed to text " + "\n")
input_array.append(input_value)
'''

input_array = []
#------------------------------------------------------------------------------------------
#Predefined list of values for testing
#------------------------------------------------------------------------------------------

uploadArray = ["12percent.mp3","criminals.mp3","dreamhouse.mp3"]

#------------------------------------------------------------------------------------------
#Writing output to transcripts
#------------------------------------------------------------------------------------------

def write_to_output(transcript):
    file = open("VoicemailText.txt","a")
    file.write(transcript)
    file.write("\n")
    file.write("\n")


#------------------------------------------------------------------------------------------
#Get transcript from Amazon server
#------------------------------------------------------------------------------------------

def get_final_transcript(url):
    text = json.load(urlopen(url))
    transcript_text = (text['results']['transcripts'][0]['transcript'])
    write_to_output(transcript_text)

#------------------------------------------------------------------------------------------
#Check status of transcription job every five seconds
#------------------------------------------------------------------------------------------

def check_job_status(job_name):
    status = (transcribe.get_transcription_job(TranscriptionJobName=job_name))
    temp = job_name
    if status['TranscriptionJob']['TranscriptionJobStatus'] == 'IN_PROGRESS':
        print ("Job Status ==> ", "IN_PROGRESS")
        time.sleep(5)
        check_job_status(temp)
    else:
        print ("Job Status ==> ", status['TranscriptionJob']['TranscriptionJobStatus'])
        transcript_url = (status['TranscriptionJob']['Transcript']['TranscriptFileUri'])
        get_final_transcript(transcript_url)

#------------------------------------------------------------------------------------------
#Send to Amazon Transcribe to get text output
#------------------------------------------------------------------------------------------

def transcribe_new_file(audiosource, audioname):
    timestampLabel = str(datetime.datetime.now().strftime("%I%M%S"))
    print ("\nTranscribing " + audiosource)
    job_name = "Transcrpt" + audioname + timestampLabel
    job_uri = audiosource
    transcribe.start_transcription_job(
        TranscriptionJobName=job_name,
        Media={'MediaFileUri': job_uri},
        MediaFormat='wav',
        LanguageCode='en-US',
        MediaSampleRateHertz=8000
    )
    check_job_status(job_name)

#------------------------------------------------------------------------------------------
#Callback for S3 upload that updates on upload progress
#------------------------------------------------------------------------------------------

class ProgressPercentage(object):
    def __init__(self, filename):
        self._filename = filename
        self._size = float(os.path.getsize(filename))
        self._seen_so_far = 0
        self._lock = threading.Lock()

    def __call__(self, bytes_amount):
        with self._lock:
            self._seen_so_far += bytes_amount
            percentage = (self._seen_so_far / self._size) * 100
            sys.stdout.write(
                "\r%s  %s / %s  (%.2f%%)" % (
                    self._filename, self._seen_so_far, self._size,
                    percentage))
            if percentage == 100:
                file_name = "https://s3.amazonaws.com/bhmoviequotes/" + self._filename
                #transcribe_new_file(file_name,self._filename)
            sys.stdout.flush()

#------------------------------------------------------------------------------------------
#Define outlook inbox and accounts from local account on computer
#------------------------------------------------------------------------------------------


outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
accounts = win32com.client.Dispatch("Outlook.Application").Session.Accounts

#------------------------------------------------------------------------------------------
#Add input values from the email attachments
#------------------------------------------------------------------------------------------

def add_input_value(value):
    input_array.append(value)
    print (input_array)

#------------------------------------------------------------------------------------------
#For all emails in all folders check if 'VM' is in the subject line and save the attachment
#------------------------------------------------------------------------------------------

def emailLoader(folder):
    messages = folder.Items
    a = len(messages)
    if a > 0:
        for tempMessage in messages:

            subject = tempMessage.Subject
            subjectStr = str(subject)
            print ("Email Subject ==> ", subjectStr)
            attachments = tempMessage.attachments
            print (attachments)
            try:
                attachment = attachments.Item(1)
                if "wav" in str(attachment):
                    print ("Email Attachment ==> ", attachment)
                    print ("Saving Attachment")
                    attachment.SaveAsFile(os.getcwd() + '\\' + str(attachment))
                    print ("Attachment Saved")
                    add_input_value(str(attachment))
            except:
                    print ("There is no attachment.")

#------------------------------------------------------------------------------------------
#Get all folders for all accounts and send to email checker
#------------------------------------------------------------------------------------------

for account in accounts:
    global inbox
    inbox = outlook.Folders(account.DeliveryStore.DisplayName)
    folders = inbox.Folders

    for folder in folders:
        if "Voicemails" in str(folder):
            emailLoader(folder)
        a = len(folder.folders)

        if a>0 :
            global z
            z = outlook.Folders(account.DeliveryStore.DisplayName).Folders(folder.name)
            x = z.Folders
            for y in x:
                emailLoader(y)

#------------------------------------------------------------------------------------------
#Run the program: Read input array, upload to bucket, send to callback
#------------------------------------------------------------------------------------------

for upload in input_array:
    if Path(upload).is_file():
        write_to_output(upload)
        print ("Uploading " + upload)
        s3Client.upload_file(upload,'bhmoviequotes',upload,Callback=ProgressPercentage(upload))
        s3.Bucket('bhmoviequotes').upload_file(upload,upload)
    else:
        print ("File does not exist.")
        time.sleep(5)
