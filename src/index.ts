// 型定義をインポート
import type {
	ChannelInfo,
	VideoItem,
	SearchResponse,
	VideosResponse,
	ChannelsResponse,
} from "./types/youtube";

// Google Apps Script services are available globally

// スプレッドシートの設定
const SPREADSHEET_ID =
	PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID") || "";
const SHEET_NAME = "YouTubeハッシュタグ分析";

// ハッシュタグのリスト
const TARGET_HASHTAGS = ["#安野たかひろ", "#チームみらい"];

// メインの処理を実行する関数
function main() {
	try {
		// スプレッドシートを取得または作成
		const spreadsheet = getOrCreateSpreadsheet();
		const sheet = getOrCreateSheet(spreadsheet, SHEET_NAME);

		// ヘッダーを設定
		setupSheetHeaders(sheet);

		// 各ハッシュタグに対して検索を実行
		for (const hashtag of TARGET_HASHTAGS) {
			searchVideosByHashtag(hashtag, sheet);
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
		"低評価数",
		"コメント数",
		"動画URL",
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
function searchVideosByHashtag(
	hashtag: string,
	sheet: GoogleAppsScript.Spreadsheet.Sheet,
) {
	Logger.log(`ハッシュタグ「${hashtag}」で動画を検索中...`);

	try {
		// YouTube Data APIを使用して動画を検索
		const searchResponse = YouTube?.Search?.list("id,snippet", {
			q: hashtag,
			type: "video",
			part: "snippet",
			maxResults: 50,
			order: "date",
			publishedAfter: new Date(
				Date.now() - 365 * 24 * 60 * 60 * 1000,
			).toISOString(),
		});

		if (!searchResponse?.items) {
			Logger.log(`No items found for hashtag: ${hashtag}`);
			return;
		}

		if (!searchResponse.items || searchResponse.items.length === 0) {
			Logger.log(`ハッシュタグ「${hashtag}」の動画は見つかりませんでした。`);
			return;
		}

		// 動画の詳細情報を取得
		const videoIds = searchResponse.items
			.map((item) => item.id?.videoId)
			.filter((id): id is string => !!id);

		if (videoIds.length === 0) {
			Logger.log(`No valid video IDs found for hashtag: ${hashtag}`);
			return;
		}

		const videosResponse = YouTube?.Videos?.list("snippet,statistics", {
			id: videoIds.join(","),
			part: "snippet,statistics",
		});

		if (!videosResponse?.items) {
			Logger.log(`No video details found for hashtag: ${hashtag}`);
			return;
		}

		if (!videosResponse.items || videosResponse.items.length === 0) {
			Logger.log("動画の詳細情報を取得できませんでした。");
			return;
		}

		// チャンネル情報を取得
		const channelIds = [
			...new Set(
				videosResponse.items
					.map((video) => video.snippet?.channelId)
					.filter((id): id is string => !!id),
			),
		];

		if (channelIds.length === 0) {
			Logger.log(`No channel IDs found for hashtag: ${hashtag}`);
			return;
		}

		const channelsResponse = YouTube?.Channels?.list("snippet,statistics", {
			id: channelIds.join(","),
			part: "snippet,statistics",
		});

		if (!channelsResponse?.items) {
			Logger.log(`No channel details found for hashtag: ${hashtag}`);
			return;
		}

		// チャンネル情報をマップに格納
		const channelInfoMap = new Map<
			string,
			{ title: string; subscriberCount: string }
		>();
		for (const channel of channelsResponse.items) {
			if (channel.id && channel.snippet) {
				channelInfoMap.set(channel.id, {
					title: channel.snippet.title || "不明",
					subscriberCount: channel.statistics?.subscriberCount || "0",
				});
			}
		}

		if (channelInfoMap.size === 0) {
			Logger.log(`No valid channel information found for hashtag: ${hashtag}`);
			return;
		}

		// スプレッドシートに書き込むデータを準備
		const now = new Date();
		const newRows: Array<Array<string | number | Date>> = [];

		for (const video of videosResponse.items) {
			if (!video.snippet) {
				Logger.log("Skipping video with missing snippet");
				continue;
			}

			const channelId = video.snippet.channelId;
			if (!channelId) {
				Logger.log(`Skipping video with missing channel ID: ${video.id}`);
				continue;
			}

			const channelInfo = channelInfoMap.get(channelId) || {
				title: "不明",
				subscriberCount: "0",
			};

			// 動画IDを取得（型安全に）
			const videoId = (() => {
				if (typeof video.id === "string") return video.id;
				if (video.id && typeof video.id === "object" && "videoId" in video.id) {
					return (video.id as { videoId: string }).videoId;
				}
				return "";
			})();

			if (!videoId) {
				Logger.log("Skipping video with missing ID");
				continue;
			}

			// 動画の統計情報を安全に取得
			const stats = video.statistics || {
				viewCount: "0",
				likeCount: "0",
				dislikeCount: "0",
				commentCount: "0",
			};

			// 動画の公開日を安全に取得
			const publishedAt = video.snippet.publishedAt;
			if (!publishedAt) {
				Logger.log(`Skipping video with missing published date: ${videoId}`);
				continue;
			}

			// 動画がショートかどうかを判定
			const isShort =
				video.snippet.title?.includes("#shorts") ||
				video.snippet.description?.includes("#shorts") ||
				video.snippet.title?.toLowerCase().includes("shorts") ||
				video.snippet.description?.toLowerCase().includes("shorts");

			newRows.push([
				now, // 取得日時
				hashtag, // ハッシュタグ
				videoId, // 動画ID
				isShort ? "ショート" : "通常", // 動画カテゴリ
				video.snippet.title || "タイトルなし", // 動画タイトル
				`https://www.youtube.com/watch?v=${videoId}`, // 動画URL
				channelInfo.title, // チャンネル名
				Number.parseInt(channelInfo.subscriberCount, 10) || 0, // チャンネル登録者数
				new Date(publishedAt), // 動画公開日
				video.snippet.description || "", // 動画の説明
				Number.parseInt(stats.viewCount || "0", 10), // 視聴回数
				Number.parseInt(stats.likeCount || "0", 10), // いいね数
				Number.parseInt(stats.dislikeCount || "0", 10), // 低評価数
				Number.parseInt(stats.commentCount || "0", 10), // コメント数
				`=HYPERLINK("https://www.youtube.com/watch?v=${videoId}", "動画を見る")`, // 動画URL（ハイパーリンク）
			]);
		}

		if (newRows.length === 0) {
			Logger.log(`No valid video data to write for hashtag: ${hashtag}`);
			return;
		}

		// スプレッドシートに追加
		if (newRows.length > 0) {
			try {
				const lastRow = sheet.getLastRow();
				const range = sheet.getRange(
					lastRow + 1,
					1,
					newRows.length,
					newRows[0].length,
				);
				range.setValues(newRows);
				Logger.log(
					`ハッシュタグ「${hashtag}」の動画を ${newRows.length} 件追加しました。`,
				);
			} catch (error) {
				const errorMessage =
					error instanceof Error ? error.message : String(error);
				Logger.log(
					`スプレッドシートへの書き込み中にエラーが発生しました: ${errorMessage}`,
				);
			}
		}
	} catch (error: unknown) {
		const errorMessage = error instanceof Error ? error.message : String(error);
		const errorStack = error instanceof Error ? error.stack : "";
		Logger.log(`エラーが発生しました: ${errorMessage}`);
		if (errorStack) {
			Logger.log(errorStack);
		}
	}
}

// 重複する動画を削除する関数（最新のデータを残す）
function removeDuplicateVideos(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
	const lastRow = sheet.getLastRow();
	if (lastRow <= 1) return; // ヘッダーのみの場合はスキップ

	const dataRange = sheet.getRange(2, 1, lastRow - 1, 14); // A:N列
	const data = dataRange.getValues();

	// 動画IDをキーとして、最新の行を保持
	const videoMap = new Map();

	data.forEach((row, index) => {
		const videoId = row[2]; // C列: 動画ID
		videoMap.set(videoId, row);
	});

	// 重複を削除したデータを作成
	const uniqueData = Array.from(videoMap.values());

	// データをクリアしてから再書き込み
	dataRange.clearContent();
	if (uniqueData.length > 0) {
		sheet
			.getRange(2, 1, uniqueData.length, uniqueData[0].length)
			.setValues(uniqueData);
	}

	const duplicateCount = data.length - uniqueData.length;
	if (duplicateCount > 0) {
		Logger.log(`重複する動画を ${duplicateCount} 件削除しました。`);
	}
}

// グローバルスコープに型をマージ
type GlobalWithMain = typeof globalThis & {
	main: () => void;
};

// 手動実行用の関数
(globalThis as GlobalWithMain).main = main;
