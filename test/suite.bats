#!/usr/bin/env bats

setup() {
  [ -d /tmp/test ] && rm -rf /tmp/test
  mkdir /tmp/test
  pushd /tmp/test
}

# We write One Big Test for everything. Two reasons:
#
# 1. Otherwise we'd need to alter suite.bats each time we add a test case
# 2. When a test fails, we want our state to allow `docker cp` (rather than
#    delete tempfiles)
@test "all tests" {
  for dir in $(find /app/test -name 'test-*'); do
    (cd /tmp/test && cat $dir/input.blob | /app/do-convert-stream-to-mime-multipart MIME-BOUNDARY "$(cat $dir/input.json)" > output.mime)
    diff --text -u /tmp/test/output.mime $dir/expect-output.mime
  done
}
