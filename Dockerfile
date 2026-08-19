FROM rocker/shiny:latest

# 1. Install system dependencies yang dibutuhkan untuk compile package R (openssl, curl, xml2, font untuk plotly)
RUN apt-get update && apt-get install -y \
    libssl-dev \
    libcurl4-openssl-dev \
    libxml2-dev \
    libfontconfig1-dev \
    libharfbuzz-dev \
    libfribidi-dev \
    libfreetype6-dev \
    libpng-dev \
    libtiff5-dev \
    libjpeg-dev \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /srv/shiny-server

COPY . /srv/shiny-server/

# 2. Install package R
RUN R -e "install.packages('shiny', repos='https://cloud.r-project.org')" && \
    R -e "install.packages('readxl', repos='https://cloud.r-project.org')" && \
    R -e "install.packages('openxlsx', repos='https://cloud.r-project.org')" && \
    R -e "install.packages('dplyr', repos='https://cloud.r-project.org')" && \
    R -e "install.packages('tidyr', repos='https://cloud.r-project.org')" && \
    R -e "install.packages('DT', repos='https://cloud.r-project.org')" && \
    R -e "install.packages('plotly', repos='https://cloud.r-project.org')" && \
    R -e "library(plotly); print(packageVersion('plotly'))"

EXPOSE 3838

# 3. Jalankan aplikasi Shiny
CMD ["R", "-e", "shiny::runApp('/srv/shiny-server', host='0.0.0.0', port=3838)"]
