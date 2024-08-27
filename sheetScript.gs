// Auto Youtube Subscription Playlist (2)
// This is a Google Apps Script that automatically adds new Youtube videos to playlists (a replacement for Youtube Collections feature).
// Code: https://github.com/Elijas/auto-youtube-subscription-playlist-2/
// Copy Spreadsheet: 
// https://docs.google.com/spreadsheets/d/1sZ9U52iuws6ijWPQTmQkXvaZSV3dZ3W9JzhnhNTX9GU/copy

// Debug flags
const debugFlag_dontUpdateTimestamp = false;
const debugFlag_dontUpdatePlaylists = false;
const debugFlag_logWhenNoNewVideosFound = false;

// Errorflags
var errorFlag = false;
var quotaExceeded = false;
var plErrorCount = 0;
var totalErrorCount = 0;

// Global values
var minLength = undefined;
var maxLength = undefined;

//
// Main Function to update all Playlists
//

function updatePlaylists(sheet) {
  // Init/get spreadsheets
  var sheetID = PropertiesService.getScriptProperties().getProperty("sheetID");
  if (!sheetID)
    onOpen();
  var spreadsheet = SpreadsheetApp.openById(sheetID);
  if (!sheet || !sheet.toString || sheet.toString() != 'Sheet')
    sheet = spreadsheet.getSheets()[0];
  if (!sheet || sheet.getRange("A3").getValue() !== "Playlist ID") {
    additional = sheet ? ", instead found sheet with name " + sheet.getName() : "";
    throw new Error("Cannot find playlist sheet, make sure the sheet with playlist IDs and channels is the first sheet (leftmost)" + additional)
  }
  var data = sheet.getDataRange().getValues();
  var debugSheet = spreadsheet.getSheetByName("DebugData");
  if (!debugSheet)
    debugSheet = spreadsheet.insertSheet("DebugData").hideSheet();
  var nextDebugCol = getNextDebugCol(debugSheet);
  var nextDebugRow = getNextDebugRow(debugSheet, nextDebugCol);
  var debugViewerSheet = spreadsheet.getSheetByName("Debug");
  initDebugEntry(debugViewerSheet, nextDebugCol, nextDebugRow);

  // For each playlist row...
  for (var row = reservedTableRows; row < sheet.getLastRow(); row++) {
    Logger.clear();
    Logger.log("Row: " + (row + 1));
    var playlistId = data[row][reservedColumnPlaylist];
    if (!playlistId || playlistId.substring(0, 1) == "#")
      continue;

    var lastTimestamp = data[row][reservedColumnTimestamp];
    if (!lastTimestamp) {
      var date = new Date();
      date.setHours(date.getHours() - DEFAULT_TIMESTAMP_HOURS);
      lastTimestamp = date.toISO8601String();
      sheet.getRange(row + 1, reservedColumnTimestamp + 1).setValue(lastTimestamp);
    }

    // Check if it's time to update
    var lastDate = new Date(lastTimestamp);
    var dateDiff = new Date() - lastDate;
    var nextTimeDiff = data[row][reservedColumnFrequency] * MILLIS_PER_HOUR;
    if (nextTimeDiff && dateDiff <= nextTimeDiff) {
      Logger.log("Skipped: Not time yet");
    } else {
      // ...get channels and playlists...
      var channelHandles = [];
      var playlistIds = [];
      for (var iColumn = reservedTableColumns; iColumn < sheet.getLastColumn(); iColumn++) {
        var sourceId = data[row][iColumn];
        // Skip empty cells
        if (!sourceId)
          continue;
        // ALL = Get all user subscriptions
        else if (sourceId == "ALL") {
          var newChannelIds = getAllChannelIds();
          if (!newChannelIds || newChannelIds.length === 0)
            addError("Could not find any subscriptions");
          else
            [].push.apply(channelHandles, newChannelIds);
        }
        // Add videos from playlists. MaybeTODO: better validation, since might interpret a channel with a name "PL..." as a playlist ID
        else if (["PL", "UU", "OL"].includes(sourceId.substring(0, 2)) && sourceId.length > 10)
          playlistIds.push(sourceId);
        // Add videos from channel upload playlists. MaybeTODO: do a better validation, since might interpret a channel with a name "UC..." as a channel ID
        else if (sourceId.substring(0, 2) == "UC" && sourceId.length > 10)
          playlistIds.push("UU" + sourceId.substring(2));
        // Add channel handles
        else if (sourceId.substring(0, 1) == "@")
          channelHandles.push(sourceId);
        // Try to find channel by username
        else {
          try {
            var user = YouTube.Channels.list('id', { forUsername: sourceId, maxResults: 1 });
            if (!user || !user.items) addError("Cannot query for user " + sourceId)
            else if (user.items.length === 0) addError("No user with name " + sourceId)
            else if (user.items.length !== 1) addError("Multiple users with name " + sourceId)
            else if (!user.items[0].id) addError("Cannot get id from user " + sourceId)
            else playlistIds.push("UU" + user.items[0].id.substring(2));
          } catch (e) {
            addError("Cannot search for channel with name " + sourceId + ", ERROR: " + "Message: [" + e.message + "] Details: " + JSON.stringify(e.details));
            checkQuotaExceeded(e);
            continue;
          }
        }
      }

      /// ...get videos from the channels...
      var newVideos = [];
      for (var i = 0; i < channelHandles.length; i++) {
        var videos = getChannelVideos(channelHandles[i], lastDate);
        if (!videos || typeof (videos) !== "object") {
          addError("Failed to get videos with channel handle " + channelHandles[i])
        } else if (debugFlag_logWhenNoNewVideosFound && videos.length === 0) {
          Logger.log("Channel with id " + channelHandles[i] + " has no new videos")
        } else {
          [].push.apply(newVideos, videos);
        }
      }
      for (var i = 0; i < playlistIds.length; i++) {
        var videos = getPlaylistVideos(playlistIds[i], lastDate)
        if (!videos || typeof (videos) !== "object") {
          addError("Failed to get videos with playlist id " + playlistIds[i])
        } else if (debugFlag_logWhenNoNewVideosFound && videos.length === 0) {
          Logger.log("Playlist with id " + playlistIds[i] + " has no new videos")
        } else {
          [].push.apply(newVideos, videos);
        }
      }

      Logger.log("Acquired " + newVideos.length + " videos")

      newVideos = applyFilters(newVideos, sheet, row);

      Logger.log("Filtering finished, left with " + newVideos.length + " videos")

      if (!errorFlag) {
        var newTimestamp = null;

        // ...add videos to playlist...
        if (!debugFlag_dontUpdatePlaylists) {
          newTimestamp = addVideosToPlaylist(playlistId, newVideos);
        } else {
          addError("Don't Update Playlists debug flag is set");
        }

        // ...delete old vidoes in playlist
        var daysBack = data[row][reservedColumnDeleteDays];
        if (daysBack && (daysBack > 0)) {
          var deleteBeforeTimestamp = new Date((new Date()).getTime() - daysBack * MILLIS_PER_DAY).toISO8601String();
          Logger.log("Delete before: " + deleteBeforeTimestamp);
          deletePlaylistItems(playlistId, deleteBeforeTimestamp);
        }

        // Default to no timestamp change in case of errors
        if (!newTimestamp)
          newTimestamp = lastTimestamp;
        // Update timestamp
        if (!debugFlag_dontUpdateTimestamp) {
          sheet.getRange(row + 1, reservedColumnTimestamp + 1).setValue(newTimestamp);
          Logger.log("Updating last update timestamp from " + lastTimestamp + " to " + newTimestamp);
        } else
          addError("Don't Update Timestamp debug flag is set. Not updating timestamp to: " + newTimestamp);
      }
    }
    Logger.log("Encountered " + plErrorCount + " errors on row " + (row + 1));

    // Print logs to DebugData sheet
    var newLogs = Logger.getLog().split("\n").slice(0, -1).map((log) => log.split(" INFO: "))
    if (newLogs.length > 0)
      debugSheet.getRange(nextDebugRow + 1, nextDebugCol + 1, newLogs.length, 2).setValues(newLogs)

    // Reset for next row
    nextDebugRow += newLogs.length;
    errorFlag = false;
    totalErrorCount += plErrorCount;
    plErrorCount = 0;
  }

  // Log finished script, only populate second column to signify end of execution when retrieving logs
  if (totalErrorCount == 0) {
    debugSheet.getRange(nextDebugRow + 1, nextDebugCol + 2).setValue("Updated all rows, script successfully finished")
  } else {
    debugSheet.getRange(nextDebugRow + 1, nextDebugCol + 2).setValue("Script did not successfully finish")
  }
  nextDebugRow += 1;

  // Clear next debug column if filled reservedDebugNumRows rows
  if (nextDebugRow > reservedDebugNumRows - 1) {
    var colIndex = 0;
    if (nextDebugCol < reservedDebugNumColumns - 2) {
      colIndex = nextDebugCol + 2;
    }
    clearDebugCol(debugSheet, colIndex)
  }
  loadLastDebugLog(debugViewerSheet);

  // Finally fail with a helpful message so that it's in the alert
  if (totalErrorCount > 0) {
    throw new Error(totalErrorCount + " errors were encountered while adding videos to playlists. Please check Debug sheet. Timestamps for respective rows " + quotaExceeded ? "have only been updated until the daily quota was reached." : "have not been updated.")
  }
}


//
// Constants
//

// Subscriptions added starting with the last day
const DEFAULT_TIMESTAMP_HOURS = 24 
const MILLIS_PER_HOUR = 3600000;
const MILLIS_PER_DAY = 86400000;
const QUOTA_EXCEEDED_REASON = "quotaExceeded";
// TODO rm this
const quotaExceededReason = "quotaExceeded";
const VIDEO_DURATION_REGEX = "P?([.,0-9]+D)?T?([.,0-9]+H)?([.,0-9]+M)?([.,0-9]+S)?";

// Reserved Row and Column indices (zero-based)
// If you use getRange remember those indices are one-based, so add + 1 in that call i.e.
// sheet.getRange(iRow + 1, reservedColumnTimestamp + 1).setValue(isodate);
const reservedTableRows = 3;        // Start of the range of the PlaylistID+ChannelID data
const reservedTableColumns = 7;     // Start of the range of the ChannelID data (0: A, 1: B, 2: C, 3: D, 4: E, 5: F, 6: G, ...)
const reservedColumnPlaylist = 0;   // Column containing playlist to add to
const reservedColumnTimestamp = 1;  // Column containing last timestamp
const reservedColumnFrequency = 2;  // Column containing number of hours until new check
const reservedColumnDeleteDays = 3; // Column containing number of days before today until videos get deleted
const reservedColumnMinSeconds = 4; // Column containing switch for using shorts filter
const reservedColumnMaxSeconds = 5; // Column containing switch for using shorts filter
// Reserved lengths
const reservedDebugNumRows = 900;   // Number of rows to use in a column before moving on to the next column in debug sheet
const reservedDebugNumColumns = 26; // Number of columns to use in debug sheet, must be at least 4 to allow infinite cycle

//
// Miscellaneous helper functions that don't require YouTube API calls
//

// Extend Date with toISO8601String with timzone support (Youtube needs ISO8601)
// https://stackoverflow.com/questions/17415579/how-to-iso-8601-format-a-date-with-timezone-offset-in-javascript
// Not to be confused with Date.toISOString which is missing timezone
Date.prototype.toISO8601String = function () {
  var tzo = -this.getTimezoneOffset(),
    dif = tzo >= 0 ? '+' : '-',
    pad = function (num) {
      var norm = Math.floor(Math.abs(num));
      return (norm < 10 ? '0' : '') + norm;
    };
  return this.getFullYear() +
    '-' + pad(this.getMonth() + 1) +
    '-' + pad(this.getDate()) +
    'T' + pad(this.getHours()) +
    ':' + pad(this.getMinutes()) +
    ':' + pad(this.getSeconds()) +
    dif + pad(tzo / 60) +
    ':' + pad(tzo % 60);
}

// Sort videos by oldest first
function sortVideos(vid1, vid2) {
  var date1 = new Date(vid1.videoPublishedAt);
  var date2 = new Date(vid2.videoPublishedAt);
  if (date1 < date2)
    return -1;
  if (date1 > date2)
    return 1;
  return 0;
}

// Converts the time components of an ISO8601 duration to seconds for comparison
// TODO this is excessive when `new Date(duration)` just works
function videoDurationToSeconds(duration) {
  const matches = duration.match(VIDEO_DURATION_REGEX);
  return parseFloat(!matches[1] ? 0 : matches[1].slice(0, -1)) * 87840 +
    parseFloat(!matches[2] ? 0 : matches[2].slice(0, -1)) * 3660 +
    parseFloat(!matches[3] ? 0 : matches[3].slice(0, -1)) * 60 +
    parseFloat(!matches[4] ? 0 : matches[4].slice(0, -1));
}

// Returns a new filtered array of videos based on the filters selected in the sheet
function applyFilters(videos, sheet, iRow) {
  var filters = []

  // Removes videos outside duration bounds if set
  var minOrMaxSet = false;
  minLength = sheet.getRange(iRow + 1, reservedColumnMinSeconds + 1).getValue();
  maxLength = sheet.getRange(iRow + 1, reservedColumnMaxSeconds + 1).getValue();
  if (minLength && minLength > 0) {
    Logger.log("Removing videos shorter than " + minLength + " seconds");
    minOrMaxSet = true;
  }
  if (maxLength && maxLength > 0) {
    Logger.log("Removing videos longer than " + maxLength + " seconds");
    minOrMaxSet = true;
  }
  if (minOrMaxSet)
    filters.push(removeVideosByDuration);

  return videos.filter(video => filters.reduce((acc, cur) => acc && cur(video.videoId), true));
}


//
// Functions to obtain channel IDs
// These DO require YouTube API calls
//

// Display dialog to get channel ID from channel name or handle
function getChannelId() {
  var ui = SpreadsheetApp.getUi();

  var result = ui.prompt(
    'Get Channel ID',
    'Please input a channel name:',
    ui.ButtonSet.OK_CANCEL);
  var button = result.getSelectedButton();
  var text = result.getResponseText();

  if (button == ui.Button.OK) {
    var results;
    if (text.substring(0, 1) == "@") {
      results = YouTube.Channels.list('id', {
        forHandle: text,
        maxResults: 1
      });
    } else {
      results = YouTube.Search.list('id', {
        q: text,
        type: "channel",
        maxResults: 50,
      });
    }

    for (var i = 0; i < results.items.length; i++) {
      var id = results.items[i].id;
      var result = ui.alert(
        'Please confirm',
        'Is this the link to the channel you want?\n\nhttps://youtube.com/channel/' + (id.channelId ?? id) + '',
        ui.ButtonSet.YES_NO);
      if (result == ui.Button.YES) {
        ui.alert('The channel ID is ' + (id.channelId ?? id));
        return;
      } else if (result == ui.Button.NO) {
        continue;
      } else {
        return;
      }
    }

    ui.alert('No results found for ' + text + '.');
  } else if (button == ui.Button.CANCEL) {
    return;
  } else if (button == ui.Button.CLOSE) {
    return;
  }
}

// Get Channel IDs from Subscriptions (ALL keyword)
function getAllChannelIds() {
  // get YT Subscriptions-List, src: https://www.reddit.com/r/youtube/comments/3br98c/a_way_to_automatically_add_subscriptions_to/
  var AboResponse, AboList = [[], []], nextPageToken = [], nptPage = 0, i, ix;
  // Workaround: nextPageToken API-Bug (this Tokens are limited to 1000 Subscriptions... but you can add more Tokens.)
  nextPageToken = ['', 'CDIQAA', 'CGQQAA', 'CJYBEAA', 'CMgBEAA', 'CPoBEAA', 'CKwCEAA', 'CN4CEAA', 'CJADEAA', 'CMIDEAA', 'CPQDEAA', 'CKYEEAA', 'CNgEEAA', 'CIoFEAA', 'CLwFEAA', 'CO4FEAA', 'CKAGEAA', 'CNIGEAA', 'CIQHEAA', 'CLYHEAA'];
  try {
    do {
      AboResponse = YouTube.Subscriptions.list('snippet', {
        mine: true,
        maxResults: 50,
        order: 'alphabetical',
        pageToken: nextPageToken[nptPage],
        fields: 'items(snippet(title,resourceId(channelId)))'
      });
      for (i = 0, ix = AboResponse.items.length; i < ix; i++) {
        AboList[0].push(AboResponse.items[i].snippet.title)
        AboList[1].push(AboResponse.items[i].snippet.resourceId.channelId)
      }
      nptPage += 1;
    } while (AboResponse.items.length > 0 && nptPage < 20);
    if (AboList[0].length !== AboList[1].length) {
      addError("While getting subscriptions, the number of titles (" + AboList[0].length + ") did not match the number of channels (" + AboList[1].length + ")."); // returns a string === error
      return []
    }
  } catch (e) {
    addError("Could not get subscribed channels, ERROR: " + "Message: [" + e.message + "] Details: " + JSON.stringify(e.details));
    checkQuotaExceeded(e);
    return [];
  }

  Logger.log('Acquired subscriptions %s', AboList[1].length);
  return AboList[1];
}

// Get video metadata from Channels but with less Quota use
// slower and date ordering is a bit messy but less quota costs
function getChannelVideos(channelHandle, startDate) {
  var uploadsPlaylistId;
  try {
    // Check Channel validity
    var results = YouTube.Channels.list('contentDetails', {
      forHandle: channelHandle
    });
    if (!results) {
      addError("YouTube channel search returned invalid response for channel with id " + channelHandle)
      return []
    } else if (!results.items || results.items.length === 0) {
      addError("Cannot find channel with id " + channelHandle)
      return []
    } else {
      uploadsPlaylistId = results.items[0].contentDetails.relatedPlaylists.uploads;
    }
  } catch (e) {
    addError("Cannot search YouTube for channel with id " + channelHandle + ", ERROR: " + "Message: [" + e.message + "] Details: " + JSON.stringify(e.details));
    checkQuotaExceeded(e);
    return [];
  }

  return getPlaylistVideos(uploadsPlaylistId, startDate);
}


//
// Functions for adding and deleting videos from playlists
// These DO require YouTube API calls
//

// Get Video metadata from Playlist
function getPlaylistVideos(playlistId, startDate) {
  var videos = [];
  var nextPageToken = '';
  while (nextPageToken != null) {
    try {
      var results = YouTube.PlaylistItems.list('contentDetails', {
        playlistId: playlistId,
        maxResults: 50,
        publishedAfter: startDate.toISO8601String(),
        pageToken: nextPageToken
      });
      if (!results || !results.items) {
        addError("YouTube playlist search returned invalid response for playlist with id " + playlistId)
        return [];
      }
    } catch (e) {
      if (e.details) {
        if (e.details.code === 404) {
          Logger.log("Warning: Channel does not have any uploads in " + playlistId + ", ignore if this is intentional as this will not fail the script. API error details for troubleshooting: " + JSON.stringify(e.details));
          return [];
        }
        Logger.log("Cannot search YouTube with playlist id " + playlistId + ", ERROR: " + "Message: [" + e.message + "] Details: " + JSON.stringify(e.details));
        // Can't do anything with the videos we already found if the quota's been exceeded
        if (checkQuotaExceeded(e)) {
          return [];
        }
      } else {
        Logger.log("Cannot search YouTube with playlist id " + playlistId + ", ERROR: " + "Message: [" + e.message + "]");
      }
      break;
    }

    // Filter out videos published before the last time this ran
    [].push.apply(videos, results.items.map(video => video.contentDetails).filter(video => startDate < new Date(video.videoPublishedAt)));
    nextPageToken = results.nextPageToken;
  }

  if (videos.length === 0) {
    try {
      // Check Playlist validity
      var results = YouTube.Playlists.list('id', {
        id: playlistId
      });
      if (!results || !results.items) {
        addError("YouTube search returned invalid response for playlist with id " + playlistId)
        return []
      } else if (results.items.length === 0) {
        addError("Cannot find playlist with id " + playlistId)
        return []
      }
    } catch (e) {
      addError("Cannot lookup playlist with id " + playlistId + " on YouTube, ERROR: " + "Message: [" + e.message + "] Details: " + JSON.stringify(e.details));
      checkQuotaExceeded(e);
      return [];
    }
  }

  return videos;
}

// Returns false if video's length is outside the min and max bounds if set
// There might be better/more accurate ways
function removeVideosByDuration(videoId) {
  try {
    var response = YouTube.Videos.list('contentDetails', {
      id: videoId,
    });
    if (response.items && response.items.length) {
      var duration = response.items[0].contentDetails.duration;
      var durationSec = videoDurationToSeconds(duration);
      // Since there can be a 1 second variation, we check for +- 1 second, due to following bug
      // https://stackoverflow.com/questions/72459082/yt-api-pulling-different-video-lengths-for-youtube-videos
      if (minLength && durationSec <= minLength + 1)
        return false;
      else if (maxLength && durationSec >= maxLength - 1)
        return false;
      else
        return true;
    }
  } catch (e) {
    addError("Problem filtering for video with id " + videoId + ", ERROR: " + "Message: [" + e.message + "]" + e.details ? " Details: " + JSON.stringify(e.details) : "");
    checkQuotaExceeded(e);
  }
  return false;
}

// Add Videos to Playlist using Video IDs and upload times
function addVideosToPlaylist(playlistId, videos) {
  var successCount = 0;
  var errorCount = 0;
  var lastSuccessTimestamp = null;
  var totalVideos = videos.length;
  videos.sort(sortVideos); // Sort videos in order to automatically handle exceeded quota
  if (0 < totalVideos) {
    for (var i = 0; i < totalVideos; i++) {
      // Use a buffer of 1 second to retry this video next time on quota exceeded failure
      lastSuccessTimestamp = new Date(new Date(videos[i].videoPublishedAt).getTime() - 1000).toISO8601String();
      try {
        YouTube.PlaylistItems.insert({
          snippet: {
            playlistId: playlistId,
            resourceId: {
              videoId: videos[i].videoId,
              kind: 'youtube#video'
            }
          }
        }, 'snippet');
        successCount++;
      } catch (e) {
        if (checkQuotaExceeded(e)) {
          // Increase error count by remaining vids (including this one)
          errorCount += totalVideos - i;
          addError("Quota exceeded while adding videos! Updating last update timestamp to pick up where this run left off: " + lastSuccessTimestamp);
          break;
        } else if (e.details.code === 409) {
          // Skip error count if Video exists in playlist already
          Logger.log("Couldn't update playlist with video (" + videos[i] + "), ERROR: Video already exists")
        } else if (e.details.code === 400 && e.details.errors[0].reason === "playlistOperationUnsupported") {
          addError("Couldn't update watch later or watch history playlist with video, functionality deprecated; try adding videos to a different playlist")
          errorCount++;
        } else {
          try {
            var results = YouTube.Videos.list('snippet', {
              id: videos[i].videoId
            });
            if (results.items.length === 0) {
              // Skip error count if video is private (found when using getPlaylistVideoIds)
              Logger.log("Couldn't update playlist with video (" + videos[i] + "), ERROR: Cannot find video, most likely private")
            } else {
              addError("Couldn't update playlist with video (" + videos[i] + "), ERROR: " + "Message: [" + e.message + "] Details: " + JSON.stringify(e.details));
              errorCount++;
            }
          } catch (e) {
            if (checkQuotaExceeded(e)) {
              // Increase error count by remaining vids (including this one)
              errorCount += totalVideos - i;
              addError("Quota exceeded while adding videos! Updating last update timestamp to pick up where this run left off: " + lastSuccessTimestamp);
              break;
            } else {
              addError("Couldn't update playlist with video (" + videos[i] + "), 404 on update, tried to search for video by id, got ERROR: Message: [" + e.message + "] Details: " + JSON.stringify(e.details));
              errorCount++;
            }
          }
        }
      }
    }
    Logger.log("Added " + successCount + " video(s) to playlist. Error for " + errorCount + " video(s).")
    errorFlag |= (errorCount > 0);
    if (!errorFlag)
      // Increase time by 1 second because Date doesn't track millis and if we don't increase the timestamp past the last vid, that vid will be readded next execution
      lastSuccessTimestamp = new Date(new Date(lastSuccessTimestamp).getTime() + 1000).toISO8601String();
  } else {
    Logger.log("No new videos yet.")
  }

  return lastSuccessTimestamp;
}

// Delete Videos from Playlist if they're older than the defined time or dupes
// If deleteBeforeTimestamp arg is uset, only delete dupes (currently unreachable code path)
function deletePlaylistItems(playlistId, deleteBeforeTimestamp = new Date(0).toISO8601String()) {
  var nextPageToken = '';
  var oldIds = [];
  var videoIdMap = new Map();
  while (nextPageToken != null) {
    try {
      var results = YouTube.PlaylistItems.list('contentDetails', {
        playlistId: playlistId,
        maxResults: 50,
        order: "date",
        pageToken: nextPageToken
      });

      results.items.forEach(video => {
        if (new Date(video.contentDetails.videoPublishedAt) < new Date(deleteBeforeTimestamp)) {
          Logger.log("Del: | " + video.contentDetails)
          oldIds.push(video.id);
        } else {
          if (!videoIdMap.has(video.contentDetails.videoId))
            videoIdMap.set(video.contentDetails.videoId, [video.id])
          else
            videoIdMap.get(video.contentDetails.videoId).push(video.id);
        }
      });

      nextPageToken = results.nextPageToken;
    } catch (e) {
      addError("Problem getting existing videos from playlist with id " + playlistId + ", ERROR: " + "Message: [" + e.message + "] Details: " + JSON.stringify(e.details));
      break;
    }
  }

  // Delete old videos
  try {
    oldIds.forEach(id => YouTube.PlaylistItems.remove(id));
  } catch (e) {
    addError("Problem deleting old videos from playlist with id " + playlistId + ", ERROR: " + "Message: [" + e.message + "] Details: " + JSON.stringify(e.details));
    checkQuotaExceeded(e);
  }

  // Delete duplicates by PlaylistItem id
  try {
    var duplicateIds = [];
    videoIdMap.forEach((ids, key, map) => [].push.apply(duplicateIds, ids.slice(1)));
    duplicateIds.forEach(id => YouTube.PlaylistItems.remove(id));
  } catch (e) {
    addError("Problem deleting duplicate videos from playlist with id " + playlistId + ", ERROR: " + "Message: [" + e.message + "] Details: " + JSON.stringify(e.details));
    checkQuotaExceeded(e);
  }
}


//
// Functions for maintaining debug logs
// These don't require YouTube API calls
//

// Log errors in debug sheet and increment error count
function addError(s) {
  Logger.log(s);
  errorFlag = true;
  plErrorCount += 1;
}

// Common function to check for quotaExceeded in caught exception
function checkQuotaExceeded(error) {
  if (error.details && error.details.errors && error.details.errors.some(
    e => e.reason == QUOTA_EXCEEDED_REASON
  )) {
    quotaExceeded = true;
    Logger.log("Daily quota exceeded");
    return true;
  }
  return false;
}

// Parse debug sheet to find column of cell to write debug logs to
function getNextDebugCol(debugSheet) {
  var data = debugSheet.getDataRange().getValues();
  // Only one column, not filled yet, return this column
  if (data.length < reservedDebugNumRows) return 0;
  // Need to iterate since next col might be in middle of data
  for (var col = 0; col < reservedDebugNumColumns; col += 2) {
    // New column
    // Necessary check since data is list of lists and col might be out of bounds
    if (data[0].length < col + 1) return col
    // Unfilled column
    if (data[reservedDebugNumRows - 1][col + 1] == "") return col;
  }
  clearDebugCol(debugSheet, 0)
  return 0;
}

// Parse debug sheet to find row of cell to write debug logs to
function getNextDebugRow(debugSheet, nextDebugCol) {
  var data = debugSheet.getDataRange().getValues();
  // Empty sheet, return first row
  if (data.length == 1 && data[0].length == 1 && data[0][0] == "") return 0;
  // Only one column, not filled yet, return last row + 1
  // Second check needed in case reservedDebugNumRows has expanded while other columns are filled
  if (data.length < reservedDebugNumRows && data[0][0] != "") return data.length;
  for (var row = 0; row < reservedDebugNumRows; row++) {
    // Found empty row
    if (data[row][nextDebugCol + 1] == "") return row;
  }
  return 0;
}

// Clear column in debug sheet for next execution's logs
function clearDebugCol(debugSheet, colIndex) {
  // Clear first reservedDebugNumRows rows
  debugSheet.getRange(1, colIndex + 1, reservedDebugNumRows, 2).clear();
  // Clear as many additional rows as necessary
  var rowIndex = reservedDebugNumRows;
  while (debugSheet.getRange(rowIndex + 1, colIndex + 1, 1, 2).getValues()[0][1] != "") {
    debugSheet.getRange(rowIndex + 1, colIndex + 1, 1, 2).clear();
    rowIndex += 1;
  }
}

// Add execution entry to debug viewer, shift previous executions and remove earliest if too many
function initDebugEntry(debugViewer, nextDebugCol, nextDebugRow) {
  // Clear currently viewing logs to get proper last row
  debugViewer.getRange("B3").clear();
  // Calculate number of existing executions
  var numExecutionsRecorded = debugViewer.getDataRange().getLastRow() - 2;
  var maxToCopy = debugViewer.getRange("B1").getValue() - 1
  var numToCopy = numExecutionsRecorded
  if (numToCopy > maxToCopy) {
    numToCopy = maxToCopy
  }
  // Shift existing executions
  debugViewer.getRange(4, 1, numToCopy, 1).setValues(debugViewer.getRange(3, 1, numToCopy, 1).getValues())
  if (numExecutionsRecorded - numToCopy > 0) {
    debugViewer.getRange(4 + numToCopy, 1, numExecutionsRecorded - numToCopy, 1).clear()
  }
  // Copy new execution
  debugViewer.getRange(3, 1).setValue("=DebugData!" + debugViewer.getRange(nextDebugRow + 1, nextDebugCol + 1).getA1Notation())
}

// Set currently viewed execution logs to most recent execution
function loadLastDebugLog(debugViewer) {
  debugViewer.getRange("B3").setValue(debugViewer.getRange("A3").getValue());
}

// Given an execution's (first log's) timestamp, return an array with the execution's logs
// Returns "" or Error if can't find logs
function getLogs(timestamp) {
  if (timestamp == "") return "";
  var debugSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DebugData");
  if (!debugSheet) throw Error("No debug logs");
  var data = debugSheet.getDataRange().getValues();
  var results = []
  for (var col = 0; col < data[0].length; col += 2) {
    for (var row = 0; row < data.length; row++) {
      if (data[row][col] == timestamp) {
        for (; row < data.length; row++) {
          if (data[row][col] == "") break;
          results.push([data[row][col + 1]]);
        }
        return results;
      }
    }
  }
  return ""
}


//
// Functions for Housekeeping
// Makes Web App, function call from Google Sheets, etc
// These don't require YouTube API calls
//

// Function to Set Up Google Spreadsheet
function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu("Youtube Controls", [
    { name: "Update Playlists", functionName: "updatePlaylists" },
    { name: "Get Channel ID", functionName: "getChannelId" }
  ]);
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheets()[0]
  if (!sheet || sheet.getRange("A3").getValue() !== "Playlist ID") {
    additional = sheet ? ", instead found sheet with name " + sheet.getName() : ""
    throw new Error("Cannot find playlist sheet, make sure the sheet with playlist IDs and channels is the first sheet (leftmost)" + additional)
  }
  PropertiesService.getScriptProperties().setProperty("sheetID", ss.getId())
}

// Function to publish Script as Web App
function doGet(e) {
  var sheetID = PropertiesService.getScriptProperties().getProperty("sheetID");
  if (e.parameter.update == "True") {
    var sheet = SpreadsheetApp.openById(sheetID).getSheets()[0];
    if (!sheet || sheet.getRange("A3").getValue() !== "Playlist ID") {
      additional = sheet ? ", instead found sheet with name " + sheet.getName() : ""
      throw new Error("Cannot find playlist sheet, make sure the sheet with playlist IDs and channels is the first sheet (leftmost)" + additional)
    }
    updatePlaylists(sheet);
  };

  var t = HtmlService.createTemplateFromFile('index.html');
  t.data = e.parameter.pl
  t.sheetID = sheetID
  return t.evaluate();
}

// Function to select playlist for Web App
function playlist(pl, sheetID) {
  var sheet = SpreadsheetApp.openById(sheetID).getSheets()[0];
  var data = sheet.getDataRange().getValues();
  if (pl == undefined) {
    pl = reservedTableRows;
  } else {
    pl = Number(pl) + reservedTableRows - 1;  // I like to think of the first playlist as being number 1.
  }
  if (pl > sheet.getLastRow()) {
    pl = sheet.getLastRow();
  }
  var playlistId = data[pl][reservedColumnPlaylist];
  return playlistId
}
