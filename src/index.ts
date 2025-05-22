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

		// 日次統計を更新
		updateDailyStats(spreadsheet);

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
		"インプレッション数",
		"インプレッションクリック率",
		"視聴回数",
		"平均再生率",
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
async function searchVideosByHashtag(
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

			// 動画の分析データを取得
			let analyticsData = null;
			try {
				analyticsData = await fetchVideoAnalytics(videoId, channelId);
			} catch (error) {
				const errorMessage =
					error instanceof Error ? error.message : String(error);
				Logger.log(
					`Error fetching analytics for video ${videoId}: ${errorMessage}`,
				);
			}

			// チャンネル登録者数の履歴を更新
			try {
				updateSubscriberHistory(
					SpreadsheetApp.openById(SPREADSHEET_ID),
					channelId,
					Number.parseInt(channelInfo.subscriberCount, 10) || 0,
				);
			} catch (error) {
				const errorMessage =
					error instanceof Error ? error.message : String(error);
				Logger.log(`Error updating subscriber history: ${errorMessage}`);
			}

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
				analyticsData?.impressions || 0, // インプレッション数
				analyticsData?.impressionsClickThroughRate
					? Number(analyticsData.impressionsClickThroughRate.toFixed(2))
					: 0, // インプレッションクリック率 (%)
				Number.parseInt(stats.viewCount || "0", 10), // 視聴回数
				analyticsData?.averageViewPercentage
					? Number(analyticsData.averageViewPercentage.toFixed(2))
					: 0, // 平均再生率 (%)
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

// 日次統計を更新する関数
function updateDailyStats(
	spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
) {
	try {
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

// 動画の分析データを取得する関数
function fetchVideoAnalytics(
	videoId: string,
	channelId: string,
): {
	averageViewPercentage?: number;
	averageViewDuration?: string;
	impressions?: number;
	impressionsClickThroughRate?: number;
	subscribersGained?: number;
	retentionRate: number[];
} | null {
	try {
		const now = new Date();
		const thirtyDaysAgo = new Date();
		thirtyDaysAgo.setDate(now.getDate() - 30);

		// 平均再生率と視聴時間を取得
		const metrics = [
			"averageViewPercentage",
			"averageViewDuration",
			"views",
			"impressions",
			"ctr",
			"subscribersGained",
		].join(",");

		// YouTube Analytics APIを使用してデータを取得
		if (!YouTubeAnalytics?.Reports) {
			Logger.log("YouTube Analytics API is not available");
			return null;
		}

		// メトリクスデータを取得
		const response = YouTubeAnalytics.Reports.query({
			ids: `channel==${channelId}`,
			startDate: thirtyDaysAgo.toISOString().split("T")[0],
			endDate: now.toISOString().split("T")[0],
			metrics: metrics,
			dimensions: "video",
			filters: `video==${videoId}`,
		});

		if (!response.rows || response.rows.length === 0) {
			Logger.log(`No analytics data found for video: ${videoId}`);
			return null;
		}

		// 視聴維持率を取得
		const retentionResponse = YouTubeAnalytics.Reports.query({
			ids: `channel==${channelId}`,
			startDate: thirtyDaysAgo.toISOString().split("T")[0],
			endDate: now.toISOString().split("T")[0],
			metrics: "audienceWatchRatio",
			dimensions: "elapsedVideoTimeRatio",
			filters: `video==${videoId}`,
			sort: "elapsedVideoTimeRatio",
		});

		const retentionRate: number[] = [];
		if (retentionResponse.rows) {
			for (const row of retentionResponse.rows) {
				const value = row[1];
				if (value !== null && value !== undefined) {
					// Ensure the value is a number before multiplying
					const numericValue =
						typeof value === "number"
							? value
							: Number.parseFloat(value as string);
					if (!Number.isNaN(numericValue)) {
						retentionRate.push(numericValue * 100); // パーセンテージに変換
					}
				}
			}
		}

		// データを整形
		const row = response.rows[0];
		const metricsIndex: Record<string, number> = {};

		if (response.columnHeaders) {
			response.columnHeaders.forEach(
				(header: { name?: string }, index: number) => {
					if (header.name) {
						metricsIndex[header.name] = index;
					}
				},
			);
		}

		const getMetric = <T>(name: string): T | undefined => {
			const index = metricsIndex[name];
			return index !== undefined ? (row[index] as T) : undefined;
		};

		return {
			averageViewPercentage: getMetric<number>("averageViewPercentage"),
			averageViewDuration: getMetric<string>("averageViewDuration"),
			impressions: getMetric<number>("impressions"),
			impressionsClickThroughRate: getMetric<number>(
				"impressionsClickThroughRate",
			),
			subscribersGained: getMetric<number>("subscribersGained"),
			retentionRate,
		};
	} catch (error) {
		const errorMessage = error instanceof Error ? error.message : String(error);
		Logger.log(
			`Error fetching analytics for video ${videoId}: ${errorMessage}`,
		);
		return null;
	}
}

// チャンネル登録者数の履歴を記録する関数
function updateSubscriberHistory(
	spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
	channelId: string,
	subscriberCount: number,
) {
	const SHEET_NAME = "チャンネル登録者数履歴";
	let sheet = spreadsheet.getSheetByName(SHEET_NAME);

	if (!sheet) {
		sheet = spreadsheet.insertSheet(SHEET_NAME);
		sheet.appendRow(["日付", "チャンネルID", "登録者数"]);
	}

	// ヘッダーが既にあるか確認
	if (sheet.getLastRow() === 0) {
		sheet.appendRow(["日付", "チャンネルID", "登録者数"]);
	}

	const now = new Date();
	sheet.appendRow([now, channelId, subscriberCount]);

	// データを日付の降順でソート
	const range = sheet.getDataRange();
	range.sort([{ column: 1, ascending: false }]);
}

// グローバルスコープに型をマージ
type GlobalWithMain = typeof globalThis & {
	main: () => void;
};

// 手動実行用の関数
(globalThis as GlobalWithMain).main = main;
