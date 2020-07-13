#!/bin/sh

# Run unit test in Windows, Linux and macOS

#----------- check for Windows, Linux and macOS build
if [ "$TRAVIS_OS_NAME" = "linux" ]
then
  # show Linux OS version
  uname -a

  # snap dependency
  sudo apt-get install -y libcanberra-gtk-module libgail-common libatk-adaptor overlay-scrollbar-gtk2 openssl

  # show available openSSL version
  ldconfig -p|grep ssl

  echo 'Start: Test SSL connection'
  #set -e
  xvfb-run -a -e /dev/stdout enduser/trackereditor -TEST_SSL
  # Check exit code
  exit_code=$?
  if [ "${exit_code}" != "0" ]
  then
	  echo "Test SSL failed: ${exit_code}"
	  exit 1
  fi
  #set +e
  echo 'Succsess: Test SSL connection'
  
  # if [ "$TRAVIS_CPU_ARCH" = "amd64" ]
  # then
  #   # Exit immediately if a command exits with a non-zero status.
  #   #set -e
  #   #xvfb-run enduser/test_trackereditor -a --format=plain
  #   #set +e
  # fi

elif [ "$TRAVIS_OS_NAME" = "osx" ]
then
  # show macOS version
  sw_vers

  # show openSSL version
  openssl version

  # Exit immediately if a command exits with a non-zero status.
  #set -e
  enduser/test_trackereditor -a --format=plain
  #set +e

elif [ "$TRAVIS_OS_NAME" = "windows" ]
then
  # Exit immediately if a command exits with a non-zero status.
  #set -e
  enduser/test_trackereditor.exe -a --format=plain
  #set +e
fi


# Remove all the extra file created by test
# We do not what it in the ZIP release files.
rm -f enduser/console_log.txt
rm -f enduser/export_trackers.txt

# Undo all changes made by testing.
git reset --hard
