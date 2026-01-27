FROM python:3.9-slim
WORKDIR /app
# Set Infineon Proxy (Crucial for internal builds)
ENV http_proxy http://proxy.infineon.com:8080
ENV https_proxy http://proxy.infineon.com:8080

# Install basic tools
RUN apt-get update && apt-get install -y curl && rm -rf /var/lib/apt/lists/*
COPY . .

# Install the libraries from the requirements.txt
RUN pip install --no-cache-dir -r requirements.txt
EXPOSE 8501

# Run the application
ENTRYPOINT ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]