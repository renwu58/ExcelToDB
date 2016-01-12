#!/bin/bash

CDIR=$(dirname $(readlink -f $0))

MAIN="com.jeffy.importexcel.ExcelToDb"

cls="$CDIR/log4j.properties"

for jar in ls $CDIR/lib/*.jar ls; do
	cls="$cls:$jar"
done

cls="$cls:$CDIR/target/importexcel-0.0.1-SNAPSHOT.jar"

echo "java -cp $cls com.jeffy.importexcel.ExcelToDb $*"

java -cp $cls com.jeffy.importexcel.ExcelToDb $*