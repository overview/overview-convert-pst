CC=gcc
CFLAGS=-I/usr/local/include/libpst-4/libpst -g
#LDFLAGS=-static -lpst -lz
LDFLAGS=-lpst -lz -lm

extract-pst: src/extract-pst.c
	$(CC) $(CFLAGS) $< -o $@ $(LDFLAGS)
