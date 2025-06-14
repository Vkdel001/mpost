# Use an official Debian-based Node.js image
FROM node:20-bullseye

# Install Python 3 and pip + required build tools
RUN apt-get update && \
    apt-get install -y python3 python3-pip python3-dev build-essential && \
    pip3 install --upgrade pip && \
    pip3 install pandas openpyxl

# Create app directory
WORKDIR /app

# Copy project files
COPY . .

# Install Node.js dependencies
RUN npm install

# Expose port (if needed)
EXPOSE 3000

# Start server
CMD ["node", "server.js"]
