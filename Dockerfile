# Use a lightweight Python image
FROM python:3.9-slim

# Set the working directory
WORKDIR /app

# Copy necessary files
COPY . /app

# Install dependencies
RUN pip install --no-cache-dir -r requirements.txt
RUN pip install gunicorn  # <-- Ensure Gunicorn is installed

# Expose the port Flask runs on
EXPOSE 5000

# Start Gunicorn
CMD ["gunicorn", "-b", "0.0.0.0:5000", "app:app"]
