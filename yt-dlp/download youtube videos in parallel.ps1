<#
    This script downloads videos from youtube playlist in parallel using yt-dlp.
    yt-dlp is a command-line media downloader supporting many platforms (such as youtube).
    https://github.com/yt-dlp/yt-dlp/

    Why this script:
        yt-dlp itself does not support downloading whole videos in parallel (it can only parallelize the download of fragments of video files).

    Notes:
    - The order of files will be the same as in the original playlist.
    - The structure of target file path is specified in the parameters -P and -o of the final yt-dlp command line. 
        E.g. these parameters: 
            -P home:$using:ytdlpHomePath -o "%(webpage_url_domain)s/%(uploader)s/$using:playlistTitle/$indexText - %(title)s.%(ext)s"
        will yield final path for file:
            'ytdlpHomePath/youtube.com/uploader/playlist name/001 - video name.(mp4,mkv,etc...)'

    Dependencies:
    - yt-dlp itself. Get it here: https://github.com/yt-dlp/yt-dlp/
    - ffmpeg - required if you use further yt-dlp functions such as file splitting, conversion, or sponsorblock removal. Get it here: https://ffmpeg.org/
#>

<#
                       _                   _   
                      (_)                 | |  
 _   _ ___  ___ _ __   _ _ __  _ __  _   _| |_ 
| | | / __|/ _ \ '__| | | '_ \| '_ \| | | | __|
| |_| \__ \  __/ |    | | | | | |_) | |_| | |_ 
 \__,_|___/\___|_|    |_|_| |_| .__/ \__,_|\__|
                              | |              
                              |_|
#>

# yt-dlp.exe must exist in this dir. This will be specific to your system.
$ytdlpRoot = "$($env:UserProfile)\totalcmd\programs\yt-dlp\"

# how many downloads to run in parallel - warning, too many will make youtube unhappy
$nrParallelDownloads = 16

# the root dir where yt-dlp will create its subdirs based on the -o parameter
# if not exists - will be created
$whereToSaveRootDir = 'C:\temp\yt-dlp downloads\'

# download this playlist - provide youtube URL
$youtubePlaylistUrl = 'https://www.youtube.com/playlist?list=PLiO4Yp-nfJ1vL4m3OuIO7EqohE3IjxmO9'

<#
                          _       
                         | |      
  _____  _____  ___ _   _| |_ ___ 
 / _ \ \/ / _ \/ __| | | | __/ _ \
|  __/>  <  __/ (__| |_| | ||  __/
 \___/_/\_\___|\___|\__,_|\__\___|
#>

Clear-Host

# add the location to the session's $PATH - this will allow to use the command-line tool straight from PS
$paths = $($env:Path -split ';') | Where-Object {$_}
If($ytdlpRoot -notin $paths) {
    $paths += $ytdlpRoot; 
    $env:Path = ($paths -join ';')
}

Write-Host "Downloading the videos from: " -NoNewline
Write-Host $youtubePlaylistUrl -ForegroundColor Magenta

# get all playlist entries as yt-dlp info JSON
Write-Host "Downloading playlist metadata"

$youtubePlaylistJson = yt-dlp --flat-playlist --dump-single-json $youtubePlaylistUrl | ConvertFrom-Json

Write-Host "Playlist metadata"
$youtubePlaylistJson | Select-Object webpage_url, title, uploader, playlist_count | Format-List | Out-String | Write-Host -ForegroundColor Magenta

Write-Host "Playlist items"
$youtubePlaylistJson.entries | Select-Object title, url | Out-String  | Write-Host -ForegroundColor Magenta

$playlistTitle = $youtubePlaylistJson.title

# test if the $whereToSaveRootDir exists - if not, create it
If(-not (Test-Path $whereToSaveRootDir)) {
    Write-Host "Creating save root directory: " -NoNewline
    Write-Host $whereToSaveRootDir -ForegroundColor Magenta
    New-Item $whereToSaveRootDir -ItemType Directory | Out-Null
}

# replace backslashes in Windows path so that it works with yt-dlp
$ytdlpHomePath = $whereToSaveRootDir.Replace('\', '/')

# create list of URLs to download in parallel, add index to allow file naming
$playlistIndex = 0
$downloadInParallelList = $youtubePlaylistJson.entries | ForEach-Object {$playlistIndex++; $_ | Select-Object @{n='index';e={$playlistIndex}}, title, url}

<#
                       _ _      _       _                     _                 _   _                   
                      | | |    | |     | |                   | |               | | | |                  
 _ __   __ _ _ __ __ _| | | ___| |   __| | _____      ___ __ | | ___   __ _  __| | | |__   ___ _ __ ___ 
| '_ \ / _` | '__/ _` | | |/ _ \ |  / _` |/ _ \ \ /\ / / '_ \| |/ _ \ / _` |/ _` | | '_ \ / _ \ '__/ _ \
| |_) | (_| | | | (_| | | |  __/ | | (_| | (_) \ V  V /| | | | | (_) | (_| | (_| | | | | |  __/ | |  __/
| .__/ \__,_|_|  \__,_|_|_|\___|_|  \__,_|\___/ \_/\_/ |_| |_|_|\___/ \__,_|\__,_| |_| |_|\___|_|  \___|
| |                                                                                                     
|_|
#>
$downloadInParallelList | ForEach-Object -ThrottleLimit $nrParallelDownloads -Parallel {
    # number format for file name - playlist file index
    $indexText = $_.index.ToString("000")
    $url = $_.URL

    # test: just write info JSON and not the file itself
    #yt-dlp --write-info-json --skip-download --no-write-thumbnail -o "%(webpage_url_domain)s/%(uploader)s/$using:playlistTitle/$indexText - %(title)s.%(ext)s" $url

    # the structure of the path of final file is specified here. There are also additional yt-dlp command-line parameters. See details in: https://github.com/yt-dlp/yt-dlp#output-template-examples
    yt-dlp --no-write-thumbnail -f "bv*[height<=1080]+ba/ba" --concurrent-fragments 4 --no-progress --output-na-placeholder "" -P home:$using:ytdlpHomePath -o "%(webpage_url_domain)s/%(uploader)s/$using:playlistTitle/$indexText - %(title)s.%(ext)s" $url
}