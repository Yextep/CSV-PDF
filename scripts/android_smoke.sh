#!/usr/bin/env bash
set -u
cd android
./gradlew --no-daemon connectedDebugAndroidTest
luma_test_status=$?
mkdir -p ../docs
if [ "$luma_test_status" -ne 0 ]; then
  adb logcat -d > ../docs/luma-android-logcat.txt
  adb shell screencap -p /sdcard/Download/luma-diagnostico.png
  adb pull /sdcard/Download/luma-diagnostico.png ../docs/luma-diagnostico.png || true
fi
for luma_photo in light dark results; do
  adb pull "/sdcard/Android/data/com.luma.app/files/luma-android-${luma_photo}.png" "../docs/luma-android-${luma_photo}.png" || true
done
exit "$luma_test_status"
