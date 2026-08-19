FROM rocker/shiny:latest

WORKDIR /srv/shiny-server

COPY . /srv/shiny-server/

RUN R -e "install.packages('shiny', repos='https://cloud.r-project.org')" && \
    R -e "install.packages('readxl', repos='https://cloud.r-project.org')" && \
    R -e "install.packages('openxlsx', repos='https://cloud.r-project.org')" && \
    R -e "install.packages('dplyr', repos='https://cloud.r-project.org')" && \
    R -e "install.packages('tidyr', repos='https://cloud.r-project.org')" && \
    R -e "install.packages('DT', repos='https://cloud.r-project.org')" && \
    R -e "install.packages('plotly', repos='https://cloud.r-project.org')" && \
    R -e "library(plotly); print(packageVersion('plotly'))"

EXPOSE 3838

CMD ["R", "-e", "shiny::runApp('/srv/shiny-server', host='0.0.0.0', port=3838)"]
