CC=gcc

## When compiling in dev mode
#CFLAGS=-I/usr/local/include/libpst-4/libpst -g -O2
#LDFLAGS=-lpst -lz -lm

## When running gdb/valgrind in Docker container:
#CFLAGS=-I/usr/local/include/libpst-4/libpst -g
#LDFLAGS=-static -lpst -lz -lm

# When compiling in Docker container:
CFLAGS=-I/usr/local/include/libpst-4/libpst -O2
LDFLAGS=-static -lpst -lz -lm -s

extract-pst: src/extract-pst.c
	$(CC) $(CFLAGS) $< -o $@ $(LDFLAGS)
