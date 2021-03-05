import times, os, strformat


let now1 = now()

when system.hostOS == "windows":
  echo now1.format("dddd")
elif system.hostOS == "linux":
  echo "running on Linux!"
elif system.hostOS == "macosx":
  echo "running on Mac OS X!"
else:
  echo "unknown operating system"


echo now1.format("dddd")



# for /f "tokens=1-4 delims=/ " %%d in ('echo %date%') do (
#     set dow=%%d
#     set month=%%e
#     set day=%%f
#     set year=%%g
# )


# :: Picking the 2th and 4th Monday
# if "%dow%"=="Mon" (
#     if %day% geq 8 if %day% leq 14 (
#         echo INFO: Today is the Second Monday
#         goto end
#     )
# )

# if "%dow%"=="Mon" (
#     if %day% geq 22 if %day% leq 28 (
#         echo INFO: Today is the Fourth Monday
#         goto end
#     )
# )

# :: add the command lines that you want to run on any other day than the 2nd and 4th Monday
# echo "neither the first nor fourth Monday of the month"