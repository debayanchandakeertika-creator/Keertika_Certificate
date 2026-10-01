# import smtplib
# from email.mime.text import MIMEText
# from email.mime.multipart import MIMEMultipart
# from email.mime.image import MIMEImage

# # Define sender and receiver email addresses
# sender_email = "" #add sender mailid
# receiver_email = "" #add reciver mail id

# # Create a message object
# message = MIMEMultipart()
# message["From"] = sender_email
# message["To"] = receiver_email
# #message["To"] = ", ".join(receiver_list)
# message["Subject"] = "Automatic Email"

# # Add body to the message
# body = "Hello, this is an automatic email!"
# message.attach(MIMEText(body, "plain"))

# # Open the image file and attach it to the message
# with open("blank1.png", "rb") as f: #add image path here
#     image_data = f.read()
#     image = MIMEImage(image_data, name="image.png")
#     message.attach(image)

# # Create a SMTP session
# smtp_session = smtplib.SMTP("smtp.office365.com", 587) #this is for outlook

# # Start TLS for security
# smtp_session.starttls()

# # Login to the email account
# smtp_session.login(sender_email, "") #add password here

# # Convert the message to a string and send it
# smtp_session.sendmail(sender_email, receiver_email, message.as_string())

# # Close the SMTP session
# smtp_session.quit()

# print("Email sent successfully!")

import numpy as np
import pandas as pd
from tensorflow.keras.models import Sequential
from tensorflow.keras.layers import Dense, LSTM, Dropout
from sklearn.model_selection import train_test_split

# Load the dataset from the CSV file
df = pd.read_csv('data.csv')

# Split the dataset into training and validation sets
train_data, val_data, train_labels, val_labels = train_test_split(df['data'].values, df['label'].values, test_size=0.2)

# Pad the sequences to a fixed length
max_length = max(len(seq) for seq in train_data)
train_data = np.array([np.pad(seq, (0, max_length - len(seq)), mode='constant') for seq in train_data])
val_data = np.array([np.pad(seq, (0, max_length - len(seq)), mode='constant') for seq in val_data])

# Convert the labels to one-hot encoding
labels = np.unique(df['label'])
num_classes = len(labels)
label_map = dict(zip(labels, range(num_classes)))
train_labels = np.array([label_map[label] for label in train_labels])
val_labels = np.array([label_map[label] for label in val_labels])
train_labels = np.eye(num_classes)[train_labels]
val_labels = np.eye(num_classes)[val_labels]

# Define the RNN model architecture
model = Sequential()
model.add(LSTM(256, input_shape=(max_length, train_data.shape[2]), return_sequences=True))
model.add(Dropout(0.2))
model.add(LSTM(128))
model.add(Dropout(0.2))
model.add(Dense(64, activation='relu'))
model.add(Dense(num_classes, activation='softmax'))

# Compile the model with appropriate loss function, optimizer, and metrics
model.compile(loss='categorical_crossentropy', optimizer='adam', metrics=['accuracy'])

# Train the model on the training data
model.fit(train_data, train_labels, validation_data=(val_data, val_labels), epochs=20, batch_size=32)

# Save the trained model to a file
model.save('rnn_model.h5')

# Load the saved model from the file
loaded_model = load_model('rnn_model.h5')

# Predict the label of a new sound file using the loaded model
new_data = # load the mel spectrogram of the new sound file as a numpy array
new_data_padded = np.pad(new_data, (0, max_length - len(new_data)), mode='constant')
new_data_padded = np.expand_dims(new_data_padded, axis=0)
prediction = loaded_model.predict(new_data_padded)
predicted_label = labels[np.argmax(prediction)]
print(f"The predicted label of the new sound file is: {predicted_label}")

