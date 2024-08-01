# Use an official Python runtime as a parent image
FROM python:3.9-slim

# Set the working directory in the container
WORKDIR /usr/src/app

# Install required packages
RUN apt-get update && apt-get install -y \
    build-essential \
    libssl-dev \
    libffi-dev \
    python3-dev \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Install Playwright and its dependencies
RUN pip install --no-cache-dir playwright
RUN playwright install-deps
RUN playwright install firefox

# Copy the current directory contents into the container at /usr/src/app
COPY . .

# Copy the Firefox profile directory into the container
COPY bd5r7ma0.default-release-temp /usr/src/app/bd5r7ma0.default-release-temp 

# Install any needed packages specified in requirements.txt
# Make sure you have a requirements.txt with playwright and any other dependencies
COPY requirements_for_decree.txt ./
RUN pip install --no-cache-dir -r requirements_for_decree.txt

# Run script.py when the container launches
CMD ["python", "./check_nghi_dinh_auto_message_v2.py"]
