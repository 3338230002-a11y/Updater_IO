FROM rocker/shiny:latest

WORKDIR /srv/shiny-server

COPY . /srv/shiny-server/

RUN R -e "install.packages(c('shiny', 'readxl', 'openxlsx', 'dplyr', 'tidyr', 'DT', 'plotly'), repos='https://cloud.r-project.org')"

EXPOSE 3838

CMD ["R", "-e", "shiny::runApp('/srv/shiny-server', host='0.0.0.0', port=3838)"]