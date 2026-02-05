# Use Infineon's internal Python image
FROM artifactory.infineon.com/docker-remote/python:3.12-slim

WORKDIR /app

# Set Infineon Proxies so we can download libraries
ENV http_proxy http://proxy.infineon.com:8080
ENV https_proxy http://proxy.infineon.com:8080

# Copy requirements and install them
COPY requirements.txt .
RUN pip install --no-cache-dir --proxy http://proxy.infineon.com:8080 -r requirements.txt

# Copy your code
COPY . .

# Expose the port HICP expects
EXPOSE 8080

# Start launcher
ENTRYPOINT ["python", "app.py"]