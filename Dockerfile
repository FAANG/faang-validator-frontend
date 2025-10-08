# Use Python 3.10 as the base image
FROM python:3.10.12

# Set working directory
WORKDIR /app

# Copy requirements file
COPY requirements.txt .

# Install dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy the entire project
COPY . .

# Set environment variables
ENV PORT=8080
ENV ENVIRONMENT=production

# Command to run the application
CMD gunicorn --bind 0.0.0.0:$PORT dash_app:server
