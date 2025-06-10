import os

def create_directory(directory):
    if os.path.exists(directory) == False:
        try:
            os.makedirs(directory)
        except Exception as e:
            print(f"Failed to create directory {directory}. Reason: {e}")