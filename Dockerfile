# Use a lightweight Python image
FROM python:3.11-slim  # Upgraded to latest Python for better performance

# Set the working directory
WORKDIR /app

# Copy necessary files
COPY . /app

# Install dependencies efficiently
RUN pip install --no-cache-dir -r requirements.txt && pip install gunicorn

# Expose the port Flask runs on
EXPOSE 5000

# Start Gunicorn with 4 worker processes for better performance
CMD ["gunicorn", "-b", "0.0.0.0:5000", "-w", "4", "app:app"]
