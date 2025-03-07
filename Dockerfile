FROM openjdk:21
LABEL authors="guilhermezuriel"

COPY target/excelTableMaker.jar app/excelTableMaker.jar
EXPOSE 8080
WORKDIR app
CMD ["java", "-jar", "excelTableMaker.jar"]