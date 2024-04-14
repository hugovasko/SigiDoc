FROM oraclelinux:7-slim

COPY . /sigidoc
WORKDIR /sigidoc

RUN ["chmod", "+x", "build.sh"]
RUN ["chmod", "+x", "sw/ant/bin/ant"]
RUN ["chmod", "+x", "scripts/create-kiln.sh"]
RUN ["chmod", "+x", "scripts/harvest-rdfs.sh"]
RUN ["chmod", "+x", "scripts/index-all.sh"]

# ADD jdk-8u401-linux-x64.tar.gz /opt //Download from https://www.oracle.com/java/technologies/downloads/#java8-linux

ENV JAVA_HOME=/opt/jdk1.8.0_401
ENV PATH=$PATH:$JAVA_HOME/bin

CMD ["./build.sh"]
