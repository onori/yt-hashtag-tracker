// 型定義をインポート
import type {
	ChannelInfo,
	VideoItem,
	SearchResponse,
	VideosResponse,
	ChannelsResponse,
} from "./types/youtube";

// Google Apps Script services are available globally

// スプレッドシートに書き込む行データの型定義
type FormattedVideoData = [
  Date,      // 取得日時
  string,    // ハッシュタグ
  string,    // 動画ID
  string,    // 動画カテゴリ ("ショート" | "通常")
  string,    // 動画タイトル
  string,    // 動画URL
  string,    // チャンネル名
  number,    // チャンネル登録者数
  Date,      // 動画公開日
  string,    // 動画の説明
  number,    // 視聴回数
  number,    // いいね数
  number     // コメント数
];

// スプレッドシートの設定
const SPREADSHEET_ID =
	PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID") || "";
const SHEET_NAME = "YouTubeハッシュタグ分析";

// ハッシュタグのリスト
const TARGET_HASHTAGS = ["#安野たかひろ", "#チームみらい"];

// メインの処理を実行する関数
async function main() {
	try {
		// スプレッドシートを取得または作成
		const spreadsheet = getOrCreateSpreadsheet();
		const sheet = getOrCreateSheet(spreadsheet, SHEET_NAME);

		// ヘッダーを設定
		setupSheetHeaders(sheet);

		// 各ハッシュタグに対して検索を実行
		for (const hashtag of TARGET_HASHTAGS) {
			await searchVideosByHashtag(hashtag, sheet);
		}

		// 重複を削除して最新のデータを残す
		removeDuplicateVideos(sheet);

		Logger.log("処理が完了しました。");
	} catch (error) {
		const errorMessage = error instanceof Error ? error.message : String(error);
		Logger.log(`Error in main function: ${errorMessage}`);
		if (error instanceof Error && error.stack) {
			Logger.log(error.stack);
		}
	}
}

// スプレッドシートを取得または作成する関数
function getOrCreateSpreadsheet() {
	if (SPREADSHEET_ID) {
		const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
		if (!spreadsheet) {
			throw new Error(`Failed to open spreadsheet with ID: ${SPREADSHEET_ID}`);
		}
		return spreadsheet;
	}

	const newSpreadsheet = SpreadsheetApp.create("YouTubeハッシュタグ分析");
	Logger.log(
		`新しいスプレッドシートが作成されました: ${newSpreadsheet.getUrl()}`,
	);
	return newSpreadsheet;
}

// シートを取得または作成する関数
function getOrCreateSheet(
	spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
	sheetName: string,
) {
	let sheet = spreadsheet.getSheetByName(sheetName);
	if (!sheet) {
		sheet = spreadsheet.insertSheet(sheetName);
		Logger.log(`新しいシートが作成されました: ${sheetName}`);
	}
	return sheet;
}

// YouTube APIから動画データを取得し整形する共通関数
async function fetchYouTubeVideoData(hashtag: string, publishedAfterISO: string): Promise<FormattedVideoData[]> {
  Logger.log(`fetchYouTubeVideoData: ハッシュタグ「${hashtag}」で動画を検索中 (公開日以降: ${publishedAfterISO})`);
  const newRows: FormattedVideoData[] = [];
  const fetchTime = new Date(); // 取得日時を一括で設定するため最初に取得

  try {
    // YouTube Data APIを使用して動画を検索
    const searchResponse = YouTube?.Search?.list("id,snippet", {
      q: hashtag,
      type: "video",
      part: "snippet",
      maxResults: 50,
      order: "date",
      publishedAfter: publishedAfterISO,
    });

    if (!searchResponse?.items || searchResponse.items.length === 0) {
      Logger.log(`fetchYouTubeVideoData: ハッシュタグ「${hashtag}」の動画は見つかりませんでした。`);
      return newRows;
    }

    const videoIds = searchResponse.items
      .map((item) => item.id?.videoId)
      .filter((id): id is string => !!id);

    if (videoIds.length === 0) {
      Logger.log(`fetchYouTubeVideoData: 有効な動画IDが見つかりませんでした: ${hashtag}`);
      return newRows;
    }

    const videosResponse = YouTube?.Videos?.list("snippet,statistics", {
      id: videoIds.join(","),
      part: "snippet,statistics",
    });

    if (!videosResponse?.items || videosResponse.items.length === 0) {
      Logger.log("fetchYouTubeVideoData: 動画の詳細情報を取得できませんでした。");
      return newRows;
    }

    const channelIds = [
      ...new Set(
        videosResponse.items
          .map((video) => video.snippet?.channelId)
          .filter((id): id is string => !!id),
      ),
    ];

    const channelInfoMap = new Map<string, { title: string; subscriberCount: string }>();
    if (channelIds.length > 0) {
      const channelsResponse = YouTube?.Channels?.list("snippet,statistics", {
        id: channelIds.join(","),
        part: "snippet,statistics",
      });
      if (channelsResponse?.items) {
        for (const channel of channelsResponse.items) {
          if (channel.id && channel.snippet) {
            channelInfoMap.set(channel.id, {
              title: channel.snippet.title || "不明",
              subscriberCount: channel.statistics?.subscriberCount || "0",
            });
          }
        }
      }
    }

    for (const video of videosResponse.items) {
      if (!video.id || !video.snippet) continue;

      const channelId = video.snippet.channelId;
      const channelInfo = channelId ? channelInfoMap.get(channelId) : null;

      const videoId = video.id;
      const stats = video.statistics || { viewCount: "0", likeCount: "0", commentCount: "0" };
      const publishedAt = video.snippet.publishedAt;

      if (!publishedAt) {
        Logger.log(`fetchYouTubeVideoData: 公開日がないためスキップ: ${videoId}`);
        continue;
      }

      const isShort =
        video.snippet.title?.includes("#shorts") ||
        video.snippet.description?.includes("#shorts") ||
        video.snippet.title?.toLowerCase().includes("shorts") ||
        video.snippet.description?.toLowerCase().includes("shorts");

      newRows.push([
        fetchTime, // 取得日時
        hashtag, // ハッシュタグ
        videoId, // 動画ID
        isShort ? "ショート" : "通常", // 動画カテゴリ
        video.snippet.title || "タイトルなし", // 動画タイトル
        `https://www.youtube.com/watch?v=${videoId}`, // 動画URL
        channelInfo?.title || "不明", // チャンネル名
        Number.parseInt(channelInfo?.subscriberCount || "0", 10) || 0, // チャンネル登録者数
        new Date(publishedAt), // 動画公開日
        video.snippet.description || "", // 動画の説明
        Number.parseInt(stats.viewCount || "0", 10) || 0, // 視聴回数
        Number.parseInt(stats.likeCount || "0", 10) || 0, // いいね数
        Number.parseInt(stats.commentCount || "0", 10) || 0, // コメント数
      ]);
    }
  } catch (error: unknown) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    Logger.log(`fetchYouTubeVideoData: エラーが発生しました (ハッシュタグ: ${hashtag}): ${errorMessage}`);
    if (error instanceof Error && error.stack) {
      Logger.log(error.stack);
    }
  }
  return newRows;
}

// シートのヘッダーを設定する関数
function setupSheetHeaders(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
	const headers = [
		"取得日時",
		"ハッシュタグ",
		"動画ID",
		"動画カテゴリ",
		"動画タイトル",
		"動画URL",
		"チャンネル名",
		"チャンネル登録者数",
		"動画公開日",
		"動画の説明",
		"視聴回数",
		"いいね数",
		"コメント数",
	];

	// ヘッダーが既に設定されているかチェック
	const range = sheet.getRange(1, 1, 1, headers.length);
	const existingHeaders = range.getValues()[0];

	if (existingHeaders[0] !== headers[0]) {
		range.setValues([headers]);
		// ヘッダー行を固定
		sheet.setFrozenRows(1);
		// ヘッダーを太字に
		range.setFontWeight("bold");
		// 列幅を自動調整
		sheet.autoResizeColumns(1, headers.length);
	}
}

// ハッシュタグで動画を検索する関数
async function searchVideosByHashtag(
	hashtag: string,
	sheet: GoogleAppsScript.Spreadsheet.Sheet,
) {
	Logger.log(`searchVideosByHashtag: ハッシュタグ「${hashtag}」で動画を検索中...`);
	try {
    const publishedAfterISO = new Date(
      Date.now() - 365 * 24 * 60 * 60 * 1000,
    ).toISOString();

    const newRows = await fetchYouTubeVideoData(hashtag, publishedAfterISO);

    if (newRows.length > 0) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
      Logger.log(
        `searchVideosByHashtag: ハッシュタグ「${hashtag}」の動画を ${newRows.length} 件追加しました。`,
      );
    } else {
      Logger.log(`searchVideosByHashtag: ハッシュタグ「${hashtag}」で追加する新しい動画はありませんでした。`);
    }
	} catch (error: unknown) {
		const errorMessage = error instanceof Error ? error.message : String(error);
		Logger.log(`searchVideosByHashtag: エラーが発生しました (ハッシュタグ: ${hashtag}): ${errorMessage}`);
    if (error instanceof Error && error.stack) {
      Logger.log(error.stack);
    }
	}
}

// 重複する動画を削除する関数（最新のデータを残す）
function removeDuplicateVideos(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
	const lastRow = sheet.getLastRow();
	if (lastRow <= 1) return; // ヘッダーのみの場合はスキップ

	// ヘッダーを除く全データを取得
	const dataRange = sheet.getDataRange();
	const data = dataRange.getValues();
	const headers = data[0]; // ヘッダー行を取得

	// データ部分のみを処理（ヘッダーを除外）
	const rows = data.slice(1);

	// 日付で降順にソート（0列目が日付）
	rows.sort((a, b) => new Date(b[0]).getTime() - new Date(a[0]).getTime());

	// 動画IDをキーとして、最新の行を保持
	const videoMap = new Map();

	for (const row of rows) {
		const videoId = row[2]; // C列: 動画ID
		if (!videoMap.has(videoId)) {
			videoMap.set(videoId, row);
		}
	}

	// 重複を削除したデータを作成（ヘッダーを先頭に追加）
	const uniqueData = [headers, ...Array.from(videoMap.values())];

	// シートをクリアしてから再書き込み
	sheet.clearContents();
	if (uniqueData.length > 0) {
		sheet
			.getRange(1, 1, uniqueData.length, uniqueData[0].length)
			.setValues(uniqueData);
	}

	const duplicateCount = rows.length - videoMap.size;
	if (duplicateCount > 0) {
		Logger.log(`重複する動画を ${duplicateCount} 件削除しました。`);
	}
}

// 日次統計を更新する関数
function updateDailyStats() {
	try {
		const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);

		const STATS_SHEET_NAME = "日次統計";
		const statsSheet = getOrCreateSheet(spreadsheet, STATS_SHEET_NAME);

		// ヘッダーを設定
		if (statsSheet.getLastRow() === 0) {
			const headers = [
				"日付",
				"ハッシュタグ",
				"動画タイプ",
				"動画数",
				"チャンネル数",
				"総再生回数",
			];
			statsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
		}

		const today = new Date();
		today.setHours(0, 0, 0, 0);
		const todayStr = Utilities.formatDate(today, "Asia/Tokyo", "yyyy/MM/dd");

		const dataSheet = spreadsheet.getSheetByName(SHEET_NAME);
		if (!dataSheet) {
			throw new Error(`シート '${SHEET_NAME}' が見つかりません`);
		}

		// データを取得（ヘッダー行を除く）
		const lastRow = dataSheet.getLastRow();
		if (lastRow <= 1) return; // データがない場合はスキップ

		const dataRange = dataSheet.getRange(2, 1, lastRow - 1, 15); // 15列分のデータを取得
		const data = dataRange.getValues();

		// 各ハッシュタグと動画タイプごとに統計を計算
		for (const hashtag of TARGET_HASHTAGS) {
			// 通常動画の統計
			const regularVideos = data.filter(
				(row) => row[1] === hashtag && row[3] === "通常",
			);

			// 通常動画の統計を常に出力（動画がなくても0で出力）
			const regularChannelCount = new Set(regularVideos.map((row) => row[6]))
				.size; // チャンネル名でユニークカウント
			const regularTotalViews = regularVideos.reduce(
				(sum, row) => sum + (Number.parseInt(row[10] || "0", 10) || 0),
				0,
			);

			statsSheet.appendRow([
				todayStr,
				hashtag,
				"通常",
				regularVideos.length,
				regularChannelCount,
				regularTotalViews,
			]);

			// ショート動画の統計
			const shortVideos = data.filter(
				(row) => row[1] === hashtag && row[3] === "ショート",
			);

			// ショート動画の統計を常に出力（動画がなくても0で出力）
			const shortChannelCount = new Set(shortVideos.map((row) => row[6])).size;
			const shortTotalViews = shortVideos.reduce(
				(sum, row) => sum + (Number.parseInt(row[10] || "0", 10) || 0),
				0,
			);

			statsSheet.appendRow([
				todayStr,
				hashtag,
				"ショート",
				shortVideos.length,
				shortChannelCount,
				shortTotalViews,
			]);
		}

		// ソート（日付の降順、ハッシュタグ、動画タイプ）
		statsSheet.getRange(2, 1, statsSheet.getLastRow() - 1, 6).sort([
			{ column: 1, ascending: false },
			{ column: 2, ascending: true },
			{ column: 3, ascending: true },
		]);

		Logger.log("日次統計を更新しました。");
	} catch (error) {
		const errorMessage = error instanceof Error ? error.message : String(error);
		Logger.log(`Error in updateDailyStats: ${errorMessage}`);
		if (error instanceof Error && error.stack) {
			Logger.log(error.stack);
		}
	}
}

// ヘッダー名から列番号を取得する関数
function getColumnIndexByHeader(
	sheet: GoogleAppsScript.Spreadsheet.Sheet,
	headerText: string,
): number {
	const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
	const index = headers.findIndex((header) => header === headerText);
	if (index === -1) {
		throw new Error(`ヘッダー "${headerText}" が見つかりませんでした`);
	}
	return index;
}

// チャンネル登録者数の履歴を記録する関数
// 毎日実行され、すべてのユニークなチャンネルの登録者数を記録する
function updateSubscriberHistory() {
	const SHEET_NAME = "チャンネル登録者数履歴";
	const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
	let sheet = spreadsheet.getSheetByName(SHEET_NAME);

	// シートが存在しない場合は作成
	if (!sheet) {
		sheet = spreadsheet.insertSheet(SHEET_NAME);
		sheet.appendRow([
			"日付",
			"チャンネルタイトル",
			"チャンネル登録者数",
			"視聴回数",
		]);
	} else if (sheet.getLastRow() === 0) {
		// シートが空の場合はヘッダーを追加
		sheet.appendRow([
			"日付",
			"チャンネルタイトル",
			"チャンネル登録者数",
			"視聴回数",
		]);
	}

	// メインのシートからデータを取得
	const mainSheet = spreadsheet.getSheetByName("YouTubeハッシュタグ分析");
	if (!mainSheet) {
		Logger.log("メインのシートが見つかりませんでした");
		return;
	}

	// ヘッダーから列番号を取得
	const channelTitleCol = getColumnIndexByHeader(mainSheet, "チャンネル名");
	const subscriberCountCol = getColumnIndexByHeader(
		mainSheet,
		"チャンネル登録者数",
	);
	const viewCountCol = getColumnIndexByHeader(mainSheet, "視聴回数");

	// データを取得（ヘッダー行を除く）
	const data = mainSheet
		.getRange(2, 1, mainSheet.getLastRow() - 1, mainSheet.getLastColumn())
		.getValues();

	// チャンネルごとの最新の登録者数と視聴回数の合計を保持するマップ
	const channelMap = new Map<
		string,
		{ title: string; count: number; viewCount: number }
	>();

	// データを処理して、各チャンネルの最新の登録者数を取得
	for (const row of data) {
		const channelTitle = row[channelTitleCol];
		const subscriberCount = row[subscriberCountCol];
		const viewCount = row[viewCountCol];

		if (
			channelTitle &&
			typeof subscriberCount === "number" &&
			!Number.isNaN(subscriberCount)
		) {
			const currentViewCount =
				typeof viewCount === "number" && !Number.isNaN(viewCount)
					? viewCount
					: 0;

			const existing = channelMap.get(channelTitle);
			if (existing) {
				// 既存のチャンネルの場合、視聴回数を加算
				// 登録者数は最新の値で更新
				if (subscriberCount > existing.count) {
					existing.count = subscriberCount;
				}
				existing.viewCount += currentViewCount;
			} else {
				// 新しいチャンネルの場合、新規に追加
				channelMap.set(channelTitle, {
					title: channelTitle,
					count: subscriberCount,
					viewCount: currentViewCount,
				});
			}
		}
	}

	// 現在の日時を取得
	const now = new Date();

	// マップのデータを行に変換
	const rows = Array.from(channelMap.entries()).map(([_, channel]) => [
		now, // 日付
		channel.title, // チャンネルタイトル
		Math.floor(channel.count), // 登録者数（整数に丸める）
		Math.floor(channel.viewCount), // 視聴回数の合計（整数に丸める）
	]);

	// データをシートに追加
	if (rows.length > 0) {
		sheet
			.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length)
			.setValues(rows);

		// データを日付の降順、チャンネルタイトルの昇順でソート
		const range = sheet.getDataRange();
		range.sort([
			{ column: 1, ascending: false }, // 日付（新しい順）
			{ column: 2, ascending: true }, // チャンネルタイトル（昇順）
		]);

		// ヘッダーを太字に
		const headerRange = sheet.getRange(1, 1, 1, 3);
		headerRange.setFontWeight("bold");

		// 列幅を自動調整
		sheet.autoResizeColumns(1, 4);

    Logger.log(`${rows.length}件のチャンネル登録者数を記録しました`);
  }
}

// ハッシュタグで動画を検索してシートに追加する関数
async function searchAndAppendVideos(hashtag: string, sheet: GoogleAppsScript.Spreadsheet.Sheet) {
  Logger.log(`searchAndAppendVideos: ハッシュタグ「${hashtag}」で動画を検索中...`);
  try {
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate()); // 今日の0時0分0秒
    const publishedAfterISO = today.toISOString();

    const newRows = await fetchYouTubeVideoData(hashtag, publishedAfterISO);

    if (newRows.length > 0) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
      Logger.log(
        `searchAndAppendVideos: ハッシュタグ「${hashtag}」の動画を ${newRows.length} 件追加しました。`,
      );
    } else {
      Logger.log(`searchAndAppendVideos: ハッシュタグ「${hashtag}」で追加する新しい動画はありませんでした。`);
    }
  } catch (error) {
    Logger.log(`searchAndAppendVideos: エラーが発生しました (ハッシュタグ: ${hashtag}): ${error}`);
    if (error instanceof Error && error.stack) {
      Logger.log(error.stack);
    }
  }
}

// 当日分のデータ内で重複を削除する関数
function removeDailyDuplicates(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const videoIdIndex = header.indexOf("動画ID");
  const dateIndex = header.indexOf("取得日時");
  
  if (videoIdIndex === -1 || dateIndex === -1) {
    Logger.log("必要なカラムが見つかりませんでした。");
    return;
  }
  
  // 当日のデータのみを抽出
  const todayData = data.filter((row, index) => {
    if (index === 0) return false; // ヘッダーをスキップ
    const rowDate = new Date(row[dateIndex]);
    return rowDate >= today;
  });
  
  if (todayData.length <= 1) {
    Logger.log("重複チェックの必要なし: 当日のデータが1件以下です。");
    return;
  }
  
  // 重複を検出（動画IDが同じで最新のものだけを残す）
  const uniqueVideos = new Map();
  const rowsToDelete: number[] = [];
  
  // データを逆順に処理して、同じ動画IDの最初の出現（最新）を保持
  for (let i = todayData.length - 1; i >= 0; i--) {
    const row = todayData[i];
    const videoId = row[videoIdIndex];
    
    if (!uniqueVideos.has(videoId)) {
      uniqueVideos.set(videoId, i);
    } else {
      // 重複している行は削除対象
      const originalRowIndex = uniqueVideos.get(videoId);
      const rowDate = new Date(row[dateIndex]);
      const originalRowDate = new Date(todayData[originalRowIndex][dateIndex]);
      
      // 日付が新しい方を保持（同じ場合は既存のものを保持）
      if (rowDate > originalRowDate) {
        rowsToDelete.push(originalRowIndex);
        uniqueVideos.set(videoId, i);
      } else {
        rowsToDelete.push(i);
      }
    }
  }
  
  // 重複する行を削除（逆順で削除する）
  if (rowsToDelete.length > 0) {
    const headerRows = 1; // ヘッダー行の数
    const firstDataRow = sheet.getFrozenRows() + 1;
    
    // 行番号を降順にソート
    const sortedRowsToDelete = [...new Set(rowsToDelete)]
      .map(i => i + firstDataRow) // 実際の行番号に変換
      .sort((a, b) => b - a); // 降順にソート
    
    // 重複行を削除
    for (const rowNum of sortedRowsToDelete) {
      sheet.deleteRow(rowNum);
    }
    
    Logger.log(`${rowsToDelete.length} 件の重複動画を削除しました。`);
  } else {
    Logger.log("重複する動画は見つかりませんでした。");
  }
}

// 日次更新を実行する関数
async function dailyUpdate() {
  try {
    // スプレッドシートを取得または作成
    const spreadsheet = getOrCreateSpreadsheet();
    const sheet = getOrCreateSheet(spreadsheet, "YouTubeハッシュタグ分析_積み上げ");

    // ヘッダーを設定
    setupSheetHeaders(sheet);

    // 各ハッシュタグに対して検索を実行し、新しい動画を追加
    for (const hashtag of TARGET_HASHTAGS) {
      await searchAndAppendVideos(hashtag, sheet);
    }

    // 当日分のデータ内で重複を削除
    removeDailyDuplicates(sheet);
    
    Logger.log("日次更新が完了しました。");
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    Logger.log(`Error in dailyUpdate: ${errorMessage}`);
    if (error instanceof Error && error.stack) {
      Logger.log(error.stack);
    }
  }
}

// グローバルスコープに型をマージ
interface GlobalWithMain {
  main: () => void;
  dailyUpdate: () => Promise<void>;
}

// 手動実行用の関数
const globalObj = globalThis as unknown as GlobalWithMain;
globalObj.main = main;
globalObj.dailyUpdate = dailyUpdate;
