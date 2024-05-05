FROM hugovasko/java8:linux-64

WORKDIR /sigidoc

COPY build.sh /sigidoc/
COPY sw/ant/bin/ant /sigidoc/sw/ant/bin/ant
COPY scripts/create-kiln.sh /sigidoc/scripts/create-kiln.sh
COPY scripts/harvest-rdfs.sh /sigidoc/scripts/harvest-rdfs.sh
COPY scripts/index-all.sh /sigidoc/scripts/index-all.sh

RUN ["chmod", "+x", "build.sh"]
RUN ["chmod", "+x", "sw/ant/bin/ant"]
RUN ["chmod", "+x", "scripts/create-kiln.sh"]
RUN ["chmod", "+x", "scripts/harvest-rdfs.sh"]
RUN ["chmod", "+x", "scripts/index-all.sh"]

CMD ["./build.sh"]
