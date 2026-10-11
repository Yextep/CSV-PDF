#!/usr/bin/env bash
set -u
cd android
./gradlew --no-daemon connectedDebugAndroidTest
luma_test_status=$?
mkdir -p ../docs
for luma_photo in light dark results; do
  adb pull "/sdcard/Android/data/com.luma.app/files/luma-android-${luma_photo}.png" "../docs/luma-android-${luma_photo}.png" || true
done
exit "$luma_test_status"
