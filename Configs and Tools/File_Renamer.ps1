# Add "Music_ " to the names of all-files:
Get-ChildItem -Path "D:/musics" -File | Rename-Item -NewName { "Music_ $($_.Name)" }

# Add "Music_ " to the names of only MP3-files:
Get-ChildItem -Path "D:/musics" -File -Filter "*.mp3" | Rename-Item -NewName { "Music_ $($_.Name)" }

# Add "Music_ " to the names of only the files whose names include the word "test":
Get-ChildItem -Path "D:/musics" -File | Where-Object { $_.Name -match "test" } | Rename-Item -NewName { "Music_ $($_.Name)" }

# Remove "Music_ " from the names of all files:
Get-ChildItem -Path "D:/musics" -File | Where-Object { $_.Name -like "Unvocal_ *" } | Rename-Item -NewName { $_.Name -replace "^Music_ ", "" }



