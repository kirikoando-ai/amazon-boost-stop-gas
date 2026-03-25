function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Boost Ops')
    .addItem('0) シート初期化', 'setupSheets')
    .addItem('0.1) config推奨値を再作成', 'setupConfigPresetsOnly')
    .addItem('0.2) configをROAS項目へ変換', 'migrateConfigAcosToRoas')
    .addItem('0.3) ROAS20入札係数をconfigへ適用', 'applyRoas20BidConfig')
    .addItem('0.4) RAWから必須inputを自動作成', 'autoBuildInputsFromRaw')
    .addItem('0.5) SP生データをinput_boostへ変換', 'convertRawSpBulkToInputBoost')
    .addItem('0.6) 外部シートのSPデータを変換', 'convertExternalSpBulkToInputBoost')
    .addSeparator()
    .addItem('1) 判定＆出力作成', 'runBoostStopWorkflow')
    .addItem('1.5) Amazon SP一括行を再生成', 'buildBulkSpSheet')
    .addItem('2) 出力シートをCSV化（Drive保存）', 'exportOutputSheetsAsCsv')
    .addToUi();
}

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheetWithHeaders_(ss, 'config', ['key', 'value', 'note']);
  ensureSheetWithHeaders_(ss, 'config_presets', ['preset', 'key', 'value', 'note', 'recommended_for']);
  ensureSheetWithHeaders_(ss, 'input_boost', [
    'boost_campaign_id',
    'boost_campaign_name',
    'boost_ad_group_id',
    'boost_ad_group_name',
    'product_key',
    'product_label',
    'target_type',
    'target_text',
    'match_type',
    'impressions',
    'clicks',
    'orders',
    'sales',
    'spend',
    'cpc',
    'target_key',
    'existing_exact_in_normal'
  ]);
  ensureSheetWithHeaders_(ss, 'input_mapping', [
    'target_key',
    'target_type',
    'target_text',
    'source_product_key',
    'source_boost_ad_group_name',
    'normal_campaign_id',
    'normal_campaign_name',
    'normal_ad_group_id',
    'normal_ad_group_name',
    'default_bid',
    'state'
  ]);
  ensureSheetWithHeaders_(ss, 'input_bulk_sp_raw', [
    'プロダクト',
    'エンティティ',
    '操作',
    'キャンペーンID',
    '広告グループID',
    'ポートフォリオID',
    '広告ID',
    'キーワードID',
    '商品ターゲティングID',
    'キャンペーン名',
    '広告グループ名',
    'キャンペーン名（情報提供のみ）',
    '広告グループ名（情報提供のみ）',
    'ポートフォリオ名（情報提供のみ）',
    '開始日',
    '終了日',
    'ターゲティングの種類',
    'ステータス',
    'キャンペーンの状態（情報提供のみ）',
    '広告グループの状態（情報提供のみ）',
    '1日の予算',
    'SKU',
    'ASIN（情報提供のみ）',
    '利用資格のステータス（情報提供のみ）',
    '対象外の理由（情報提供のみ）',
    '広告グループの入札額の初期値',
    '広告グループの入札額の初期値（情報提供のみ）',
    '入札額',
    'キーワードテキスト',
    'ネイティブ言語のキーワード',
    'ネイティブ言語の地域',
    'マッチタイプ',
    '入札戦略',
    '掲載枠',
    '割合',
    '商品ターゲティング式',
    '解決済みの商品ターゲティング式（情報提供のみ）',
    'オーディエンスID',
    '購入者のコホート割合',
    'お客様コホートタイプ',
    'セグメント名（情報提供のみ）',
    'インプレッション数',
    'クリック数',
    'クリックスルー率',
    '支出',
    '売上',
    '注文',
    '商品点数',
    'コンバージョン率',
    'ACOS',
    'CPC',
    'ROAS'
  ]);
  ensureSheetWithHeaders_(ss, 'input_boost_campaigns', [
    'boost_campaign_id',
    'boost_campaign_name_contains',
    'product_key',
    'product_label',
    'enabled'
  ]);
  ensureSheetWithHeaders_(ss, 'input_stop_products', [
    'product_key',
    'product_label',
    'approved_to_stop',
    'note'
  ]);
  ensureSheetWithHeaders_(ss, 'output_migrate_exact', [
    'action',
    'entity',
    'operation',
    'state',
    'campaign_id',
    'campaign_name',
    'ad_group_id',
    'ad_group_name',
    'target_type',
    'target_text',
    'match_type',
    'bid',
    'reason',
    'source_boost_campaign_id',
    'source_boost_ad_group_id',
    'clicks',
    'orders',
    'roas',
    'cvr'
  ]);
  ensureSheetWithHeaders_(ss, 'output_pause_boost', [
    'action',
    'entity',
    'operation',
    'state',
    'boost_campaign_id',
    'boost_campaign_name',
    'boost_ad_group_id',
    'boost_ad_group_name',
    'pause_level',
    'reason',
    'unit_impressions'
  ]);
  ensureSheetWithHeaders_(ss, 'output_blockers', [
    'blocker_type',
    'boost_campaign_id',
    'boost_campaign_name',
    'boost_ad_group_id',
    'boost_ad_group_name',
    'target_type',
    'target_text',
    'reason',
    'required_action'
  ]);
  ensureSheetWithHeaders_(ss, 'output_summary', ['metric', 'value']);
  ensureSheetWithHeaders_(ss, 'output_bulk_sp', [
    'プロダクト',
    'エンティティ',
    '操作',
    'キャンペーンID',
    '広告グループID',
    'ポートフォリオID',
    '広告ID',
    'キーワードID',
    '商品ターゲティングID',
    'キャンペーン名',
    '広告グループ名',
    '開始日',
    '終了日',
    'ターゲティングの種類',
    'ステータス',
    '1日の予算',
    'SKU',
    '広告グループの入札額の初期値',
    '入札額',
    'キーワードテキスト',
    'ネイティブ言語のキーワード',
    'ネイティブ言語の地域',
    'マッチタイプ',
    '入札戦略',
    '掲載枠',
    '割合',
    '商品ターゲティング式',
    'オーディエンスID',
    '購入者のコホート割合',
    'お客様コホートタイプ'
  ]);

  writeDefaultConfig_(ss.getSheetByName('config'));
  writeConfigPresets_(ss.getSheetByName('config_presets'));
  SpreadsheetApp.getUi().alert('シート初期化が完了しました。config_presets で保守/標準/攻めを確認し、config に採用値を入力してから「判定＆出力作成」を実行してください。');
}

function setupConfigPresetsOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheetWithHeaders_(ss, 'config_presets', ['preset', 'key', 'value', 'note', 'recommended_for']);
  writeConfigPresets_(ss.getSheetByName('config_presets'));
  SpreadsheetApp.getUi().alert('config_presets を更新しました。config は変更していません。');
}

function migrateConfigAcosToRoas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('config');
  if (!sheet) {
    throw new Error('config シートがありません。先に「0) シート初期化」を実行してください。');
  }

  const rows = getRows_(sheet);
  const notes = {
    min_impressions_to_stop: '停止対象にする最低表示回数（Boost側の停止単位合計）',
    min_clicks_to_migrate: '移行判定の最低クリック数',
    min_orders_to_migrate: '移行判定の最低注文数',
    min_roas_to_migrate: '移行判定の最低ROAS',
    min_cvr_to_migrate: '移行判定の最低CVR(%)',
    fallback_bid: 'CPC・default_bid が無いときの入札額',
    use_roas_bid_multiplier: 'true なら ROAS帯に応じて入札係数を適用',
    roas_bid_multiplier_ge_35: 'ROAS>=35 の入札係数',
    roas_bid_multiplier_ge_25: '25<=ROAS<35 の入札係数',
    roas_bid_multiplier_ge_20: '20<=ROAS<25 の入札係数',
    roas_bid_multiplier_ge_15: '15<=ROAS<20 の入札係数',
    roas_bid_multiplier_lt_15: 'ROAS<15 の入札係数',
    min_bid: '移行入札の下限',
    max_bid: '移行入札の上限',
    bulk_product_type: '一括アップロードの Product 列 (例: Sponsored Products)',
    boost_campaign_name_contains: '生データ変換時のBoost判定用キャンペーン名キーワード（カンマ区切り）',
    require_stop_approval: 'true なら input_stop_products で停止承認済み product_key のみ処理',
    source_spreadsheet_id: '外部取り込み元スプレッドシートID（0.6利用時）',
    source_sp_sheet_name: '外部取り込み元のSPシート名（0.6利用時）',
    pause_level: 'ad_group または campaign'
  };

  const map = {};
  rows.forEach(function(r) {
    const key = String(r.key || '').trim();
    if (key) {
      map[key] = String(r.value || '').trim();
    }
  });

  if (!map.min_roas_to_migrate) {
    const oldAcos = toNumber_(map.max_acos_to_migrate);
    if (oldAcos > 0) {
      map.min_roas_to_migrate = round2_(100 / oldAcos);
    } else {
      map.min_roas_to_migrate = defaultConfig_().min_roas_to_migrate;
    }
  }
  delete map.max_acos_to_migrate;

  const keyOrder = [
    'min_impressions_to_stop',
    'min_clicks_to_migrate',
    'min_orders_to_migrate',
    'min_roas_to_migrate',
    'min_cvr_to_migrate',
    'fallback_bid',
    'use_roas_bid_multiplier',
    'roas_bid_multiplier_ge_35',
    'roas_bid_multiplier_ge_25',
    'roas_bid_multiplier_ge_20',
    'roas_bid_multiplier_ge_15',
    'roas_bid_multiplier_lt_15',
    'min_bid',
    'max_bid',
    'bulk_product_type',
    'boost_campaign_name_contains',
    'require_stop_approval',
    'source_spreadsheet_id',
    'source_sp_sheet_name',
    'pause_level'
  ];

  const output = keyOrder.map(function(key) {
    return [key, map[key] != null ? map[key] : defaultConfig_()[key], notes[key] || ''];
  });

  sheet.clearContents();
  sheet.getRange(1, 1, 1, 3).setValues([['key', 'value', 'note']]);
  sheet.getRange(2, 1, output.length, 3).setValues(output);

  SpreadsheetApp.getUi().alert('config を ROAS 項目へ変換しました。min_roas_to_migrate を確認してください。');
}

function applyRoas20BidConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('config');
  if (!sheet) {
    throw new Error('config シートがありません。先に「0) シート初期化」を実行してください。');
  }

  const values = {
    use_roas_bid_multiplier: true,
    roas_bid_multiplier_ge_35: 0.85,
    roas_bid_multiplier_ge_25: 0.75,
    roas_bid_multiplier_ge_20: 0.65,
    roas_bid_multiplier_ge_15: 0.55,
    roas_bid_multiplier_lt_15: 0.45,
    min_bid: 20,
    max_bid: 120
  };

  const notes = {
    use_roas_bid_multiplier: 'true なら ROAS帯に応じて入札係数を適用',
    roas_bid_multiplier_ge_35: 'ROAS>=35 の入札係数',
    roas_bid_multiplier_ge_25: '25<=ROAS<35 の入札係数',
    roas_bid_multiplier_ge_20: '20<=ROAS<25 の入札係数',
    roas_bid_multiplier_ge_15: '15<=ROAS<20 の入札係数',
    roas_bid_multiplier_lt_15: 'ROAS<15 の入札係数',
    min_bid: '移行入札の下限',
    max_bid: '移行入札の上限'
  };

  upsertConfigRows_(sheet, values, notes);
  SpreadsheetApp.getUi().alert('ROAS20向け入札係数を config に適用しました。');
}

function convertRawSpBulkToInputBoost() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = loadConfig_(ss.getSheetByName('config'));
  const rawRows = getRows_(ss.getSheetByName('input_bulk_sp_raw'));
  const boostCampaignRows = getRows_(ss.getSheetByName('input_boost_campaigns'));
  const boostCampaignMap = buildBoostCampaignMap_(boostCampaignRows);
  const asinIndex = buildAsinIndexFromRaw_(rawRows);

  if (!rawRows.length) {
    throw new Error('input_bulk_sp_raw にデータがありません。SPバルクの2枚目シートを貼り付けてください。');
  }

  const converted = convertSpRowsToInputBoost_(rawRows, boostCampaignMap, config, asinIndex);

  const inputBoostSheet = ss.getSheetByName('input_boost');
  writeRows_(inputBoostSheet, converted);

  SpreadsheetApp.getUi().alert(
    [
      'input_boost 変換が完了しました。',
      '変換行数: ' + converted.length,
      '対象: SP + キーワード/商品ターゲティング + Boost判定一致'
    ].join('\n')
  );
}

function autoBuildInputsFromRaw() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = loadConfig_(ss.getSheetByName('config'));
  const rawRows = getRows_(ss.getSheetByName('input_bulk_sp_raw'));
  if (!rawRows.length) {
    throw new Error('input_bulk_sp_raw にデータがありません。先にRAWを取り込んでください。');
  }

  const existingBoostCampaignRows = getRows_(ss.getSheetByName('input_boost_campaigns'));

  const asinIndex = buildAsinIndexFromRaw_(rawRows);
  const autoBoostCampaignRows = buildBoostCampaignRowsFromRaw_(rawRows, config, existingBoostCampaignRows, asinIndex);
  writeRows_(ss.getSheetByName('input_boost_campaigns'), autoBoostCampaignRows);

  ensureSheetHasColumns_(ss.getSheetByName('input_mapping'), ['source_product_key', 'source_boost_ad_group_name']);
  const autoMappingRows = buildMappingRowsFromRaw_(rawRows, config, autoBoostCampaignRows);
  writeRows_(ss.getSheetByName('input_mapping'), autoMappingRows);

  const boostCampaignMap = buildBoostCampaignMap_(autoBoostCampaignRows);
  const convertedBoost = convertSpRowsToInputBoost_(rawRows, boostCampaignMap, config, asinIndex);
  writeRows_(ss.getSheetByName('input_boost'), convertedBoost);

  SpreadsheetApp.getUi().alert(
    [
      'RAWから必須inputの自動作成が完了しました。',
      'input_boost_campaigns: ' + autoBoostCampaignRows.length + ' 行',
      'input_mapping: ' + autoMappingRows.length + ' 行',
      'input_boost: ' + convertedBoost.length + ' 行',
      '注意: input_stop_products は自動更新していません（停止ASINは手入力）。',
      '次: input_stop_products に停止ASINを入力して approved_to_stop=true を付け、1) 判定＆出力作成'
    ].join('\n')
  );
}

function convertExternalSpBulkToInputBoost() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = loadConfig_(ss.getSheetByName('config'));
  const boostCampaignRows = getRows_(ss.getSheetByName('input_boost_campaigns'));
  const boostCampaignMap = buildBoostCampaignMap_(boostCampaignRows);

  if (!config.source_spreadsheet_id) {
    throw new Error('config.source_spreadsheet_id が未設定です。外部データ元スプレッドシートIDを設定してください。');
  }

  const sourceSs = SpreadsheetApp.openById(config.source_spreadsheet_id);
  const sourceSheet = sourceSs.getSheetByName(config.source_sp_sheet_name);
  if (!sourceSheet) {
    throw new Error('外部シートが見つかりません: ' + config.source_sp_sheet_name);
  }

  const sourceRows = getRows_(sourceSheet);
  if (!sourceRows.length) {
    throw new Error('外部シートにデータがありません。');
  }

  const asinIndex = buildAsinIndexFromRaw_(sourceRows);
  const converted = convertSpRowsToInputBoost_(sourceRows, boostCampaignMap, config, asinIndex);

  writeRows_(ss.getSheetByName('input_boost'), converted);

  SpreadsheetApp.getUi().alert(
    [
      '外部シートから input_boost 変換が完了しました。',
      '元シート: ' + config.source_sp_sheet_name,
      '変換行数: ' + converted.length
    ].join('\n')
  );
}

function convertSpRowsToInputBoost_(rawRows, boostCampaignMap, config, asinIndex) {
  const converted = [];

  rawRows.forEach(function(r) {
    if (String(r['プロダクト'] || '').trim() !== 'スポンサープロダクト広告') {
      return;
    }

    const entity = String(r['エンティティ'] || '').trim();
    const targetType = normalizeTargetTypeFromEntity_(entity);
    if (!targetType) {
      return;
    }

    const campaignId = String(r['キャンペーンID'] || '').trim();
    const campaignName = String(r['キャンペーン名'] || r['キャンペーン名（情報提供のみ）'] || '').trim();
    const boostInfo = resolveBoostCampaign_(campaignId, campaignName, boostCampaignMap, config);
    if (!boostInfo.isBoost) {
      return;
    }

    const targetText = targetType === 'keyword'
      ? String(r['キーワードテキスト'] || '').trim()
      : String(r['商品ターゲティング式'] || r['解決済みの商品ターゲティング式（情報提供のみ）'] || '').trim();
    if (!targetText) {
      return;
    }

    const asin = String(r['ASIN（情報提供のみ）'] || '').trim();
    const sku = String(r['SKU'] || '').trim();
    const adGroupId = String(r['広告グループID'] || '').trim();
    const adGroupAsin = asinIndex ? String(asinIndex.byAdGroup[campaignId + '|' + adGroupId] || '').trim() : '';
    const campaignAsin = asinIndex ? String(asinIndex.byCampaign[campaignId] || '').trim() : '';
    const productKey = asin || boostInfo.productKey || adGroupAsin || campaignAsin || sku;

    converted.push({
      boost_campaign_id: campaignId,
      boost_campaign_name: campaignName,
      boost_ad_group_id: String(r['広告グループID'] || '').trim(),
      boost_ad_group_name: String(r['広告グループ名'] || r['広告グループ名（情報提供のみ）'] || '').trim(),
      product_key: productKey,
      product_label: boostInfo.productLabel || String(r['SKU'] || '').trim(),
      target_type: targetType,
      target_text: targetText,
      match_type: normalizeMatchType_(String(r['マッチタイプ'] || '').trim()),
      impressions: toNumber_(r['インプレッション数']),
      clicks: toNumber_(r['クリック数']),
      orders: toNumber_(r['注文']),
      sales: toNumber_(r['売上']),
      spend: toNumber_(r['支出']),
      cpc: toNumber_(r['CPC']),
      target_key: '',
      existing_exact_in_normal: ''
    });
  });

  return converted;
}

function runBoostStopWorkflow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = loadConfig_(ss.getSheetByName('config'));
  const boostRows = getRows_(ss.getSheetByName('input_boost'));
  const mapRows = getRows_(ss.getSheetByName('input_mapping'));
  const stopProductRows = getRows_(ss.getSheetByName('input_stop_products'));
  const rawRows = getRows_(ss.getSheetByName('input_bulk_sp_raw'));

  if (!boostRows.length) {
    throw new Error('input_boost にデータがありません。');
  }

  const mapping = buildMapping_(mapRows);
  const approvedStopKeys = loadApprovedStopKeys_(stopProductRows);
  const approvedUnitsFromRaw = buildApprovedUnitsFromRaw_(rawRows, approvedStopKeys, config.pause_level);
  const result = evaluateRows_(boostRows, mapping, config, approvedStopKeys, approvedUnitsFromRaw);

  writeRows_(ss.getSheetByName('output_migrate_exact'), result.migrations);
  writeRows_(ss.getSheetByName('output_pause_boost'), result.pauses);
  writeRows_(ss.getSheetByName('output_blockers'), result.blockers);
  writeSummary_(ss.getSheetByName('output_summary'), result.summary);
  writeRows_(ss.getSheetByName('output_bulk_sp'), buildBulkSpRows_(result.migrations, result.pauses, config));

  SpreadsheetApp.getUi().alert(
    [
      '判定が完了しました。',
      '移行行: ' + result.summary.migration_rows,
      'アーカイブ行: ' + result.summary.pause_rows,
      'ブロッカー: ' + result.summary.blocker_rows,
      'SP一括行: ' + result.summary.bulk_sp_rows
    ].join('\n')
  );
}

function buildBulkSpSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = loadConfig_(ss.getSheetByName('config'));
  const migrations = getRows_(ss.getSheetByName('output_migrate_exact'));
  const pauses = getRows_(ss.getSheetByName('output_pause_boost'));
  const bulkRows = buildBulkSpRows_(migrations, pauses, config);
  writeRows_(ss.getSheetByName('output_bulk_sp'), bulkRows);
  SpreadsheetApp.getUi().alert('output_bulk_sp を再生成しました。行数: ' + bulkRows.length);
}

function exportOutputSheetsAsCsv() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targets = ['output_migrate_exact', 'output_pause_boost', 'output_blockers', 'output_summary', 'output_bulk_sp'];
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const parent = DriveApp.getFileById(ss.getId()).getParents().hasNext()
    ? DriveApp.getFileById(ss.getId()).getParents().next()
    : DriveApp.getRootFolder();

  const created = [];
  targets.forEach(function(name) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      return;
    }
    const values = sheet.getDataRange().getDisplayValues();
    if (!values.length) {
      return;
    }
    const csv = values
      .map(function(row) {
        return row
          .map(function(cell) {
            const safe = String(cell == null ? '' : cell).replace(/"/g, '""');
            return '"' + safe + '"';
          })
          .join(',');
      })
      .join('\n');

    const file = parent.createFile(name + '_' + timestamp + '.csv', csv, MimeType.CSV);
    created.push(file.getName());
  });

  SpreadsheetApp.getUi().alert(
    created.length
      ? 'CSVを作成しました:\n' + created.join('\n')
      : '出力対象シートにデータがありません。'
  );
}

function evaluateRows_(boostRows, mapping, config, approvedStopKeys, approvedUnitsFromRaw) {
  const migrations = [];
  const blockers = [];
  const unitMap = {};
  const migrationDedup = {};
  const approvedUnits = buildApprovedUnits_(boostRows, approvedStopKeys, config.pause_level);
  const mergedApprovedUnits = mergeApprovedUnits_(approvedUnits, approvedUnitsFromRaw || {});
  let skippedRowsByStopControl = 0;

  boostRows.forEach(function(r) {
    const unitKey = makePauseUnitKey_(r, config.pause_level);
    const stopControl = shouldProcessRowByStopApproval_(r, unitKey, config, approvedStopKeys, mergedApprovedUnits);
    if (!stopControl.allowed) {
      skippedRowsByStopControl += 1;
      if (stopControl.reason === 'missing_product_key') {
        blockers.push({
          blocker_type: 'missing_product_key_for_stop_control',
          boost_campaign_id: r.boost_campaign_id,
          boost_campaign_name: r.boost_campaign_name,
          boost_ad_group_id: r.boost_ad_group_id,
          boost_ad_group_name: r.boost_ad_group_name,
          target_type: r.target_type,
          target_text: r.target_text,
          reason: '停止承認制御が有効だが product_key が空',
          required_action: 'input_boost に product_key を設定するか、config.require_stop_approval を false に変更'
        });
      }
      return;
    }

    if (!unitMap[unitKey]) {
      unitMap[unitKey] = {
        row: r,
        impressions: 0,
        exposureEligible: false
      };
    }

    const impressions = toNumber_(r.impressions);
    const clicks = toNumber_(r.clicks);
    const orders = toNumber_(r.orders);
    const sales = toNumber_(r.sales);
    const spend = toNumber_(r.spend);
    const cpc = toNumber_(r.cpc);

    const roas = spend > 0 ? (sales / spend) : 0;
    const cvr = clicks > 0 ? (orders / clicks) * 100 : 0;

    unitMap[unitKey].impressions += impressions;
    // 停止承認制御が有効な場合は、input_stop_products 承認済み時点で停止対象にする。
    // 露出条件はクライアント判断を優先し、システム側ではゲートにしない。
    if (config.require_stop_approval) {
      unitMap[unitKey].exposureEligible = true;
    } else if (unitMap[unitKey].impressions >= config.min_impressions_to_stop) {
      unitMap[unitKey].exposureEligible = true;
    }

    // product_targeting は移行不要（停止判定のみ対象）
    if (String(r.target_type || '') === 'product_targeting') {
      return;
    }

    const shouldMigrate =
      clicks >= config.min_clicks_to_migrate &&
      orders >= config.min_orders_to_migrate &&
      roas >= config.min_roas_to_migrate &&
      cvr >= config.min_cvr_to_migrate;

    if (!shouldMigrate) {
      return;
    }

    const alreadyExists = toBoolean_(r.existing_exact_in_normal);
    if (alreadyExists) {
      return;
    }

    const key = getTargetKey_(r);
    const destination = resolveMappingDestination_(mapping, key, r);

    if (!destination) {
      blockers.push({
        blocker_type: 'migration_required_but_mapping_missing',
        boost_campaign_id: r.boost_campaign_id,
        boost_campaign_name: r.boost_campaign_name,
        boost_ad_group_id: r.boost_ad_group_id,
        boost_ad_group_name: r.boost_ad_group_name,
        target_type: r.target_type,
        target_text: r.target_text,
        reason: '好調キーワードだが移行先が見つからない',
        required_action: 'input_mapping を確認（Manualのみ / boost除外 / 同名広告グループ）'
      });
      return;
    }

    const dedupKey = [
      destination.normal_campaign_id,
      destination.normal_ad_group_id,
      r.target_type,
      r.target_text,
      'exact'
    ].join('|');

    if (migrationDedup[dedupKey]) {
      return;
    }
    migrationDedup[dedupKey] = true;

    const defaultBid = toNumber_(destination.default_bid);
    const baseBid = defaultBid > 0 ? defaultBid : (cpc > 0 ? cpc : config.fallback_bid);
    const bid = calcMigrationBid_(baseBid, roas, config);

    migrations.push({
      action: 'add_to_normal',
      entity: r.target_type === 'keyword' ? 'Keyword' : 'ProductTargeting',
      operation: 'create',
      state: destination.state || 'enabled',
      campaign_id: destination.normal_campaign_id,
      campaign_name: destination.normal_campaign_name,
      ad_group_id: destination.normal_ad_group_id,
      ad_group_name: destination.normal_ad_group_name,
      target_type: r.target_type,
      target_text: r.target_text,
      match_type: 'exact',
      bid: round2_(bid),
      reason: 'Boost停止前の好調ターゲット移行',
      source_boost_campaign_id: r.boost_campaign_id,
      source_boost_ad_group_id: r.boost_ad_group_id,
      clicks: clicks,
      orders: orders,
      roas: round2_(roas),
      cvr: round2_(cvr)
    });
  });

  const pauses = [];
  Object.keys(unitMap).forEach(function(unitKey) {
    const unit = unitMap[unitKey];
    if (!unit.exposureEligible) {
      return;
    }

    pauses.push({
      action: 'archive_boost',
      entity: config.pause_level === 'campaign' ? 'Campaign' : 'AdGroup',
      operation: 'update',
      state: 'archived',
      boost_campaign_id: unit.row.boost_campaign_id,
      boost_campaign_name: unit.row.boost_campaign_name,
      boost_ad_group_id: unit.row.boost_ad_group_id,
      boost_ad_group_name: unit.row.boost_ad_group_name,
      pause_level: config.pause_level,
      reason: '露出確認済みのためBoostをアーカイブ',
      unit_impressions: round2_(unit.impressions)
    });
  });

  return {
    migrations: migrations,
    pauses: pauses,
    blockers: blockers,
    summary: {
      boost_rows: boostRows.length,
      approved_product_keys: Object.keys(approvedStopKeys).length,
      approved_units: Object.keys(mergedApprovedUnits).length,
      skipped_rows_by_stop_control: skippedRowsByStopControl,
      migration_rows: migrations.length,
      pause_rows: pauses.length,
      blocker_rows: blockers.length,
      pause_level: config.pause_level,
      min_impressions_to_stop: config.min_impressions_to_stop,
      min_clicks_to_migrate: config.min_clicks_to_migrate,
      min_orders_to_migrate: config.min_orders_to_migrate,
      min_roas_to_migrate: config.min_roas_to_migrate,
      min_cvr_to_migrate: config.min_cvr_to_migrate,
      bulk_sp_rows: migrations.length + pauses.length
    }
  };
}

function loadApprovedStopKeys_(rows) {
  const keys = {};
  rows.forEach(function(r) {
    const key = normalizeProductKey_(r.product_key);
    if (!key) {
      return;
    }
    if (toBoolean_(r.approved_to_stop)) {
      keys[key] = true;
    }
  });
  return keys;
}

function shouldProcessRowByStopApproval_(row, unitKey, config, approvedStopKeys, approvedUnits) {
  if (!config.require_stop_approval) {
    return { allowed: true, reason: 'approval_not_required' };
  }
  const key = normalizeProductKey_(row.product_key);
  if (key && approvedStopKeys[key]) {
    return { allowed: true, reason: 'approved_product' };
  }
  if (approvedUnits[unitKey]) {
    return { allowed: true, reason: 'approved_unit_by_variation' };
  }
  return { allowed: false, reason: key ? 'not_approved_product' : 'missing_product_key' };
}

function buildApprovedUnits_(boostRows, approvedStopKeys, pauseLevel) {
  const units = {};
  boostRows.forEach(function(r) {
    const key = normalizeProductKey_(r.product_key);
    if (!key || !approvedStopKeys[key]) {
      return;
    }
    const unitKey = makePauseUnitKey_(r, pauseLevel);
    units[unitKey] = true;
  });
  return units;
}

function buildApprovedUnitsFromRaw_(rawRows, approvedStopKeys, pauseLevel) {
  const units = {};
  if (!rawRows || !rawRows.length) {
    return units;
  }
  rawRows.forEach(function(r) {
    if (String(r['プロダクト'] || '').trim() !== 'スポンサープロダクト広告') {
      return;
    }
    const asin = normalizeProductKey_(r['ASIN（情報提供のみ）']);
    if (!asin || !approvedStopKeys[asin]) {
      return;
    }
    const campaignId = String(r['キャンペーンID'] || '').trim();
    if (!campaignId) {
      return;
    }
    if (pauseLevel === 'campaign') {
      units[campaignId + '|campaign'] = true;
      return;
    }
    const adGroupId = String(r['広告グループID'] || '').trim();
    if (!adGroupId) {
      return;
    }
    units[campaignId + '|' + adGroupId] = true;
  });
  return units;
}

function mergeApprovedUnits_(unitsA, unitsB) {
  const merged = {};
  Object.keys(unitsA || {}).forEach(function(k) { merged[k] = true; });
  Object.keys(unitsB || {}).forEach(function(k) { merged[k] = true; });
  return merged;
}

function buildBulkSpRows_(migrations, pauses, config) {
  const rows = [];

  migrations.forEach(function(m) {
    const row = blankSpBulkRow_();
    row['プロダクト'] = config.bulk_product_type;
    row['エンティティ'] = m.target_type === 'keyword' ? 'Keyword' : 'Product Targeting';
    row['操作'] = 'Create';
    row['キャンペーンID'] = m.campaign_id;
    row['広告グループID'] = m.ad_group_id;
    row['キャンペーン名'] = m.campaign_name;
    row['広告グループ名'] = m.ad_group_name;
    row['ステータス'] = normalizeBulkState_(m.state || 'enabled');
    row['入札額'] = toNumber_(m.bid);
    row['キーワードテキスト'] = m.target_type === 'keyword' ? m.target_text : '';
    row['マッチタイプ'] = m.target_type === 'keyword' ? 'exact' : '';
    row['商品ターゲティング式'] = m.target_type === 'keyword' ? '' : m.target_text;
    rows.push(row);
  });

  pauses.forEach(function(p) {
    const row = blankSpBulkRow_();
    row['プロダクト'] = config.bulk_product_type;
    row['エンティティ'] = p.pause_level === 'campaign' ? 'Campaign' : 'Ad Group';
    row['操作'] = 'Update';
    row['キャンペーンID'] = p.boost_campaign_id;
    row['広告グループID'] = p.pause_level === 'campaign' ? '' : p.boost_ad_group_id;
    row['キャンペーン名'] = p.boost_campaign_name;
    row['広告グループ名'] = p.pause_level === 'campaign' ? '' : p.boost_ad_group_name;
    row['ステータス'] = normalizeBulkState_(p.state || 'archived');
    rows.push(row);
  });

  return rows;
}

function blankSpBulkRow_() {
  return {
    'プロダクト': '',
    'エンティティ': '',
    '操作': '',
    'キャンペーンID': '',
    '広告グループID': '',
    'ポートフォリオID': '',
    '広告ID': '',
    'キーワードID': '',
    '商品ターゲティングID': '',
    'キャンペーン名': '',
    '広告グループ名': '',
    '開始日': '',
    '終了日': '',
    'ターゲティングの種類': '',
    'ステータス': '',
    '1日の予算': '',
    'SKU': '',
    '広告グループの入札額の初期値': '',
    '入札額': '',
    'キーワードテキスト': '',
    'ネイティブ言語のキーワード': '',
    'ネイティブ言語の地域': '',
    'マッチタイプ': '',
    '入札戦略': '',
    '掲載枠': '',
    '割合': '',
    '商品ターゲティング式': '',
    'オーディエンスID': '',
    '購入者のコホート割合': '',
    'お客様コホートタイプ': ''
  };
}

function normalizeBulkState_(state) {
  const s = String(state || '').toLowerCase().trim();
  if (s === 'enabled' || s === 'enable') {
    return 'enabled';
  }
  if (s === 'paused' || s === 'pause') {
    return 'paused';
  }
  if (s === 'archived' || s === 'archive') {
    return 'archived';
  }
  return 'enabled';
}

function buildMapping_(rows) {
  const byTarget = {};
  const byGroupProduct = {};
  rows.forEach(function(r) {
    if (!isAllowedMappingDestination_(r)) {
      return;
    }
    const key = getTargetKey_(r);
    if (key) {
      if (!byTarget[key]) {
        byTarget[key] = [];
      }
      byTarget[key].push(r);
    }

    const sourceProduct = String(r.source_product_key || '').trim();
    const sourceAdGroup = String(r.source_boost_ad_group_name || '').trim();
    if (sourceProduct && sourceAdGroup) {
      const groupKey = sourceProduct + '|' + normalizeAdGroupNameForMatch_(sourceAdGroup);
      if (!byGroupProduct[groupKey]) {
        byGroupProduct[groupKey] = [];
      }
      byGroupProduct[groupKey].push(r);
    }
  });
  return { byTarget: byTarget, byGroupProduct: byGroupProduct };
}

function buildBoostCampaignRowsFromRaw_(rawRows, config, existingRows, asinIndex) {
  const byId = {};

  existingRows.forEach(function(r) {
    const id = String(r.boost_campaign_id || '').trim();
    if (!id) return;
    byId[id] = {
      boost_campaign_id: id,
      boost_campaign_name_contains: String(r.boost_campaign_name_contains || '').trim(),
      product_key: String(r.product_key || '').trim(),
      product_label: String(r.product_label || '').trim(),
      enabled: String(r.enabled || 'true').trim() || 'true'
    };
  });

  rawRows.forEach(function(r) {
    if (String(r['プロダクト'] || '').trim() !== 'スポンサープロダクト広告') {
      return;
    }
    const campaignId = String(r['キャンペーンID'] || '').trim();
    const campaignName = String(r['キャンペーン名'] || r['キャンペーン名（情報提供のみ）'] || '').trim();
    if (!campaignId) return;

    const isBoost = resolveBoostCampaign_(
      campaignId,
      campaignName,
      { byId: {}, containsList: [] },
      config
    ).isBoost;
    if (!isBoost) return;

    const asin = String(r['ASIN（情報提供のみ）'] || '').trim();
    const sku = String(r['SKU'] || '').trim();
    const campaignAsin = asinIndex ? String(asinIndex.byCampaign[campaignId] || '').trim() : '';
    const productKey = asin || campaignAsin || sku;
    const productLabel = String(r['SKU'] || '').trim();
    if (!byId[campaignId]) {
      byId[campaignId] = {
        boost_campaign_id: campaignId,
        boost_campaign_name_contains: '',
        product_key: productKey,
        product_label: productLabel,
        enabled: 'true'
      };
    } else {
      if ((!byId[campaignId].product_key || looksLikeNumericId_(byId[campaignId].product_key)) && asin) {
        byId[campaignId].product_key = asin;
      } else if (!byId[campaignId].product_key && productKey) {
        byId[campaignId].product_key = productKey;
      }
      if (!byId[campaignId].product_label && productLabel) byId[campaignId].product_label = productLabel;
    }
  });

  return Object.keys(byId)
    .sort()
    .map(function(id) { return byId[id]; });
}

function buildMappingRowsFromRaw_(rawRows, config, boostCampaignRows) {
  const boostMap = buildBoostCampaignMap_(boostCampaignRows);
  const normalDestByGroupProduct = {};
  const seen = {};
  const rows = [];
  const asinIndex = buildAsinIndexFromRaw_(rawRows);

  // 1) 通常側の移行先候補を構築（Manualのみ / boost・全商品・B2B除外）
  rawRows.forEach(function(r) {
    if (String(r['プロダクト'] || '').trim() !== 'スポンサープロダクト広告') return;
    const campaignId = String(r['キャンペーンID'] || '').trim();
    const campaignName = String(r['キャンペーン名'] || r['キャンペーン名（情報提供のみ）'] || '').trim();
    const boostInfo = resolveBoostCampaign_(campaignId, campaignName, boostMap, config);
    if (boostInfo.isBoost) return;
    if (isExcludedDestinationCampaignName_(campaignName)) return;
    if (!isManualCampaignRow_(r)) return;

    const productKey = getRowProductKeyFromRaw_(r, asinIndex, boostInfo.productKey);
    const normalAdGroupName = String(r['広告グループ名'] || r['広告グループ名（情報提供のみ）'] || '').trim();
    if (!productKey || !normalAdGroupName) return;
    const gk = productKey + '|' + normalizeAdGroupNameForMatch_(normalAdGroupName);
    if (!normalDestByGroupProduct[gk]) {
      normalDestByGroupProduct[gk] = {
        normal_campaign_id: campaignId,
        normal_campaign_name: campaignName,
        normal_ad_group_id: String(r['広告グループID'] || '').trim(),
        normal_ad_group_name: normalAdGroupName,
        default_bid: toNumber_(r['入札額'] || r['広告グループの入札額の初期値'] || ''),
        state: 'enabled'
      };
    }
  });

  // 2) Boost側キーワードを起点にマッピングを作成（同一ASIN×同名広告グループ）
  rawRows.forEach(function(r) {
    if (String(r['プロダクト'] || '').trim() !== 'スポンサープロダクト広告') return;

    const entity = String(r['エンティティ'] || '').trim();
    const targetType = normalizeTargetTypeFromEntity_(entity);
    if (targetType !== 'keyword') return;

    const campaignId = String(r['キャンペーンID'] || '').trim();
    const campaignName = String(r['キャンペーン名'] || r['キャンペーン名（情報提供のみ）'] || '').trim();
    const boostInfo = resolveBoostCampaign_(campaignId, campaignName, boostMap, config);
    if (!boostInfo.isBoost) return;

    const targetText = String(r['キーワードテキスト'] || '').trim();
    if (!targetText) return;

    const productKey = getRowProductKeyFromRaw_(r, asinIndex, boostInfo.productKey);
    const boostAdGroupName = String(r['広告グループ名'] || r['広告グループ名（情報提供のみ）'] || '').trim();
    if (!productKey || !boostAdGroupName) return;
    const gk = productKey + '|' + normalizeAdGroupNameForMatch_(boostAdGroupName);
    const dest = normalDestByGroupProduct[gk];
    if (!dest) return;

    const targetKey = targetType + '::' + targetText.toLowerCase();
    const dedupKey = [targetKey, productKey, boostAdGroupName, dest.normal_campaign_id, dest.normal_ad_group_id].join('|');
    if (seen[dedupKey]) return;
    seen[dedupKey] = true;

    rows.push({
      target_key: targetKey,
      target_type: targetType,
      target_text: targetText,
      source_product_key: productKey,
      source_boost_ad_group_name: boostAdGroupName,
      normal_campaign_id: dest.normal_campaign_id,
      normal_campaign_name: dest.normal_campaign_name,
      normal_ad_group_id: dest.normal_ad_group_id,
      normal_ad_group_name: dest.normal_ad_group_name,
      default_bid: dest.default_bid,
      state: dest.state
    });
  });

  return rows;
}

function looksLikeNumericId_(value) {
  return /^\d{6,}$/.test(String(value || '').trim());
}

function getRowProductKeyFromRaw_(r, asinIndex, fallback) {
  const campaignId = String(r['キャンペーンID'] || '').trim();
  const adGroupId = String(r['広告グループID'] || '').trim();
  const asin = String(r['ASIN（情報提供のみ）'] || '').trim();
  const adGroupAsin = asinIndex ? String(asinIndex.byAdGroup[campaignId + '|' + adGroupId] || '').trim() : '';
  const campaignAsin = asinIndex ? String(asinIndex.byCampaign[campaignId] || '').trim() : '';
  const sku = String(r['SKU'] || '').trim();
  return asin || adGroupAsin || campaignAsin || String(fallback || '').trim() || sku;
}

function isExcludedDestinationCampaignName_(campaignName) {
  const name = String(campaignName || '').toLowerCase();
  return name.indexOf('boost') !== -1 || name.indexOf('全商品') !== -1 || name.indexOf('b2b') !== -1;
}

function isManualCampaignRow_(r) {
  const t = String(r['ターゲティングの種類'] || '').toLowerCase();
  if (!t) {
    return true;
  }
  return t.indexOf('manual') !== -1 || t.indexOf('マニュアル') !== -1;
}

function isAutoCampaignName_(campaignName) {
  const name = String(campaignName || '').toLowerCase();
  return name.indexOf('auto') !== -1 || name.indexOf('オート') !== -1;
}

function isAllowedMappingDestination_(row) {
  return !isAutoCampaignName_(row.normal_campaign_name);
}

function normalizeAdGroupNameForMatch_(name) {
  let s = String(name || '').trim().toUpperCase();
  s = s.replace(/\s+/g, '').replace(/[-_]/g, '');
  if (/^[A-Z0-9]+[A-Z]$/.test(s)) {
    s = s.slice(0, -1);
  }
  return s;
}

function adGroupNamesRoughlyMatch_(a, b) {
  const na = normalizeAdGroupNameForMatch_(a);
  const nb = normalizeAdGroupNameForMatch_(b);
  return !!na && !!nb && na === nb;
}

function adGroupSetHasMatch_(adGroupNames, targetName) {
  return adGroupNames.some(function(name) {
    return adGroupNamesRoughlyMatch_(name, targetName);
  });
}

function resolveMappingDestination_(mapping, targetKey, boostRow) {
  const byTarget = (mapping.byTarget[targetKey] || []).filter(isAllowedMappingDestination_);
  if (byTarget.length === 1) {
    return byTarget[0];
  }

  const groupKey = String(boostRow.product_key || '').trim() + '|' + normalizeAdGroupNameForMatch_(String(boostRow.boost_ad_group_name || '').trim());
  const byGroup = (mapping.byGroupProduct[groupKey] || []).filter(isAllowedMappingDestination_);
  if (byGroup.length === 1) {
    return byGroup[0];
  }

  if (byTarget.length > 1) {
    const exactSource = byTarget.filter(function(r) {
      return String(r.source_product_key || '').trim() === String(boostRow.product_key || '').trim() &&
        adGroupNamesRoughlyMatch_(String(r.source_boost_ad_group_name || '').trim(), String(boostRow.boost_ad_group_name || '').trim());
    });
    if (exactSource.length === 1) {
      return exactSource[0];
    }
  }

  return null;
}

function ensureSheetHasColumns_(sheet, requiredHeaders) {
  if (!sheet) return;
  const lastCol = Math.max(1, sheet.getLastColumn());
  const current = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const normalized = current.map(function(v) { return String(v || '').trim(); });
  const missing = requiredHeaders.filter(function(h) { return normalized.indexOf(h) === -1; });
  if (!missing.length) return;
  const startCol = normalized.length + 1;
  sheet.getRange(1, startCol, 1, missing.length).setValues([missing]);
}

function buildAsinIndexFromRaw_(rawRows) {
  const byCampaign = {};
  const byAdGroup = {};

  rawRows.forEach(function(r) {
    if (String(r['プロダクト'] || '').trim() !== 'スポンサープロダクト広告') {
      return;
    }
    const asin = String(r['ASIN（情報提供のみ）'] || '').trim();
    if (!asin) {
      return;
    }
    const campaignId = String(r['キャンペーンID'] || '').trim();
    const adGroupId = String(r['広告グループID'] || '').trim();
    if (campaignId && !byCampaign[campaignId]) {
      byCampaign[campaignId] = asin;
    }
    if (campaignId && adGroupId) {
      const key = campaignId + '|' + adGroupId;
      if (!byAdGroup[key]) {
        byAdGroup[key] = asin;
      }
    }
  });

  return { byCampaign: byCampaign, byAdGroup: byAdGroup };
}

function buildBoostCampaignMap_(rows) {
  const byId = {};
  const containsList = [];
  rows.forEach(function(r) {
    const enabled = String(r.enabled || '').trim();
    if (enabled && !toBoolean_(enabled)) {
      return;
    }
    const id = String(r.boost_campaign_id || '').trim();
    const contains = String(r.boost_campaign_name_contains || '').trim();
    const info = {
      productKey: String(r.product_key || '').trim(),
      productLabel: String(r.product_label || '').trim()
    };
    if (id) {
      byId[id] = info;
    }
    if (contains) {
      containsList.push({
        contains: contains.toLowerCase(),
        info: info
      });
    }
  });
  return {
    byId: byId,
    containsList: containsList
  };
}

function resolveBoostCampaign_(campaignId, campaignName, boostCampaignMap, config) {
  if (campaignId && boostCampaignMap.byId[campaignId]) {
    return {
      isBoost: true,
      productKey: boostCampaignMap.byId[campaignId].productKey,
      productLabel: boostCampaignMap.byId[campaignId].productLabel
    };
  }

  const nameLower = String(campaignName || '').toLowerCase();
  for (let i = 0; i < boostCampaignMap.containsList.length; i += 1) {
    const c = boostCampaignMap.containsList[i];
    if (nameLower.indexOf(c.contains) !== -1) {
      return {
        isBoost: true,
        productKey: c.info.productKey,
        productLabel: c.info.productLabel
      };
    }
  }

  const tokens = config.boost_campaign_name_contains_tokens;
  for (let j = 0; j < tokens.length; j += 1) {
    if (nameLower.indexOf(tokens[j]) !== -1) {
      return {
        isBoost: true,
        productKey: '',
        productLabel: ''
      };
    }
  }

  return {
    isBoost: false,
    productKey: '',
    productLabel: ''
  };
}

function normalizeTargetTypeFromEntity_(entity) {
  if (entity === 'キーワード') {
    return 'keyword';
  }
  if (entity === '商品ターゲティング') {
    return 'product_targeting';
  }
  return '';
}

function normalizeMatchType_(value) {
  const v = String(value || '').toLowerCase();
  if (!v) {
    return '';
  }
  if (v.indexOf('完全') !== -1 || v.indexOf('exact') !== -1) {
    return 'exact';
  }
  if (v.indexOf('フレーズ') !== -1 || v.indexOf('phrase') !== -1) {
    return 'phrase';
  }
  if (v.indexOf('部分') !== -1 || v.indexOf('broad') !== -1) {
    return 'broad';
  }
  return value;
}

function getTargetKey_(row) {
  const explicit = String(row.target_key || '').trim();
  if (explicit) {
    return explicit;
  }
  const type = String(row.target_type || '').trim().toLowerCase();
  const text = String(row.target_text || '').trim().toLowerCase();
  if (!type || !text) {
    return '';
  }
  return type + '::' + text;
}

function makePauseUnitKey_(row, pauseLevel) {
  if (pauseLevel === 'campaign') {
    return String(row.boost_campaign_id || '') + '|campaign';
  }
  return String(row.boost_campaign_id || '') + '|' + String(row.boost_ad_group_id || '');
}

function loadConfig_(sheet) {
  const defaults = defaultConfig_();
  const rows = getRows_(sheet);
  const config = Object.assign({}, defaults);

  rows.forEach(function(r) {
    const key = String(r.key || '').trim();
    if (!key) {
      return;
    }
    config[key] = r.value;
  });

  return {
    min_impressions_to_stop: toNumber_(config.min_impressions_to_stop),
    min_clicks_to_migrate: toNumber_(config.min_clicks_to_migrate),
    min_orders_to_migrate: toNumber_(config.min_orders_to_migrate),
    min_roas_to_migrate: toNumber_(
      config.min_roas_to_migrate ||
      (toNumber_(config.max_acos_to_migrate) > 0 ? 100 / toNumber_(config.max_acos_to_migrate) : '')
    ),
    min_cvr_to_migrate: toNumber_(config.min_cvr_to_migrate),
    fallback_bid: toNumber_(config.fallback_bid),
    use_roas_bid_multiplier: toBoolean_(config.use_roas_bid_multiplier),
    roas_bid_multiplier_ge_35: toNumber_(config.roas_bid_multiplier_ge_35),
    roas_bid_multiplier_ge_25: toNumber_(config.roas_bid_multiplier_ge_25),
    roas_bid_multiplier_ge_20: toNumber_(config.roas_bid_multiplier_ge_20),
    roas_bid_multiplier_ge_15: toNumber_(config.roas_bid_multiplier_ge_15),
    roas_bid_multiplier_lt_15: toNumber_(config.roas_bid_multiplier_lt_15),
    min_bid: toNumber_(config.min_bid),
    max_bid: toNumber_(config.max_bid),
    bulk_product_type: String(config.bulk_product_type || 'Sponsored Products').trim(),
    boost_campaign_name_contains_tokens: String(config.boost_campaign_name_contains || 'boost,ブースト')
      .split(',')
      .map(function(v) {
        return v.trim().toLowerCase();
      })
      .filter(function(v) {
        return v !== '';
      }),
    require_stop_approval: toBoolean_(config.require_stop_approval),
    source_spreadsheet_id: String(config.source_spreadsheet_id || '').trim(),
    source_sp_sheet_name: String(config.source_sp_sheet_name || 'スポンサープロダクト広告キャンペーン').trim(),
    pause_level: String(config.pause_level || 'ad_group').toLowerCase() === 'campaign' ? 'campaign' : 'ad_group'
  };
}

function defaultConfig_() {
  return {
    min_impressions_to_stop: 3000,
    min_clicks_to_migrate: 8,
    min_orders_to_migrate: 1,
    min_roas_to_migrate: 2.86,
    min_cvr_to_migrate: 8,
    fallback_bid: 45,
    use_roas_bid_multiplier: true,
    roas_bid_multiplier_ge_35: 0.85,
    roas_bid_multiplier_ge_25: 0.75,
    roas_bid_multiplier_ge_20: 0.65,
    roas_bid_multiplier_ge_15: 0.55,
    roas_bid_multiplier_lt_15: 0.45,
    min_bid: 20,
    max_bid: 120,
    bulk_product_type: 'Sponsored Products',
    boost_campaign_name_contains: 'boost,ブースト',
    require_stop_approval: true,
    source_spreadsheet_id: '',
    source_sp_sheet_name: 'スポンサープロダクト広告キャンペーン',
    pause_level: 'ad_group'
  };
}

function writeDefaultConfig_(sheet) {
  const defaults = defaultConfig_();
  const notes = {
    min_impressions_to_stop: '停止対象にする最低表示回数（Boost側の停止単位合計）',
    min_clicks_to_migrate: '移行判定の最低クリック数',
    min_orders_to_migrate: '移行判定の最低注文数',
    min_roas_to_migrate: '移行判定の最低ROAS',
    min_cvr_to_migrate: '移行判定の最低CVR(%)',
    fallback_bid: 'CPC・default_bid が無いときの入札額',
    use_roas_bid_multiplier: 'true なら ROAS帯に応じて入札係数を適用',
    roas_bid_multiplier_ge_35: 'ROAS>=35 の入札係数',
    roas_bid_multiplier_ge_25: '25<=ROAS<35 の入札係数',
    roas_bid_multiplier_ge_20: '20<=ROAS<25 の入札係数',
    roas_bid_multiplier_ge_15: '15<=ROAS<20 の入札係数',
    roas_bid_multiplier_lt_15: 'ROAS<15 の入札係数',
    min_bid: '移行入札の下限',
    max_bid: '移行入札の上限',
    bulk_product_type: '一括アップロードの Product 列 (例: Sponsored Products)',
    boost_campaign_name_contains: '生データ変換時のBoost判定用キャンペーン名キーワード（カンマ区切り）',
    require_stop_approval: 'true なら input_stop_products で停止承認済み product_key のみ処理',
    source_spreadsheet_id: '外部取り込み元スプレッドシートID（0.6利用時）',
    source_sp_sheet_name: '外部取り込み元のSPシート名（0.6利用時）',
    pause_level: 'ad_group または campaign'
  };

  sheet.clearContents();
  sheet.getRange(1, 1, 1, 3).setValues([['key', 'value', 'note']]);
  const rows = Object.keys(defaults).map(function(k) {
    return [k, defaults[k], notes[k] || ''];
  });
  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
}

function writeConfigPresets_(sheet) {
  const notes = {
    min_impressions_to_stop: '停止対象にする最低表示回数（Boost側の停止単位合計）',
    min_clicks_to_migrate: '移行判定の最低クリック数',
    min_orders_to_migrate: '移行判定の最低注文数',
    min_roas_to_migrate: '移行判定の最低ROAS',
    min_cvr_to_migrate: '移行判定の最低CVR(%)',
    fallback_bid: 'CPC・default_bid が無いときの入札額',
    use_roas_bid_multiplier: 'true なら ROAS帯に応じて入札係数を適用',
    roas_bid_multiplier_ge_35: 'ROAS>=35 の入札係数',
    roas_bid_multiplier_ge_25: '25<=ROAS<35 の入札係数',
    roas_bid_multiplier_ge_20: '20<=ROAS<25 の入札係数',
    roas_bid_multiplier_ge_15: '15<=ROAS<20 の入札係数',
    roas_bid_multiplier_lt_15: 'ROAS<15 の入札係数',
    min_bid: '移行入札の下限',
    max_bid: '移行入札の上限',
    pause_level: 'ad_group または campaign',
    require_stop_approval: 'true なら input_stop_products 承認済みのみ処理'
  };

  const presets = [
    {
      preset: 'conservative',
      recommended_for: '誤移行を最小化したい / 実績安定を優先',
      values: {
        min_impressions_to_stop: 5000,
        min_clicks_to_migrate: 15,
        min_orders_to_migrate: 2,
        min_roas_to_migrate: 3.33,
        min_cvr_to_migrate: 10,
        fallback_bid: 40,
        use_roas_bid_multiplier: true,
        roas_bid_multiplier_ge_35: 0.8,
        roas_bid_multiplier_ge_25: 0.7,
        roas_bid_multiplier_ge_20: 0.6,
        roas_bid_multiplier_ge_15: 0.5,
        roas_bid_multiplier_lt_15: 0.4,
        min_bid: 20,
        max_bid: 110,
        pause_level: 'ad_group',
        require_stop_approval: true
      }
    },
    {
      preset: 'standard',
      recommended_for: '現行の標準運用',
      values: {
        min_impressions_to_stop: 3000,
        min_clicks_to_migrate: 8,
        min_orders_to_migrate: 1,
        min_roas_to_migrate: 2.86,
        min_cvr_to_migrate: 8,
        fallback_bid: 45,
        use_roas_bid_multiplier: true,
        roas_bid_multiplier_ge_35: 0.85,
        roas_bid_multiplier_ge_25: 0.75,
        roas_bid_multiplier_ge_20: 0.65,
        roas_bid_multiplier_ge_15: 0.55,
        roas_bid_multiplier_lt_15: 0.45,
        min_bid: 20,
        max_bid: 120,
        pause_level: 'ad_group',
        require_stop_approval: true
      }
    },
    {
      preset: 'aggressive',
      recommended_for: '露出拡大を優先 / 早めに移行したい',
      values: {
        min_impressions_to_stop: 2000,
        min_clicks_to_migrate: 5,
        min_orders_to_migrate: 1,
        min_roas_to_migrate: 2.22,
        min_cvr_to_migrate: 5,
        fallback_bid: 50,
        use_roas_bid_multiplier: true,
        roas_bid_multiplier_ge_35: 0.9,
        roas_bid_multiplier_ge_25: 0.8,
        roas_bid_multiplier_ge_20: 0.7,
        roas_bid_multiplier_ge_15: 0.6,
        roas_bid_multiplier_lt_15: 0.5,
        min_bid: 20,
        max_bid: 130,
        pause_level: 'ad_group',
        require_stop_approval: true
      }
    }
  ];

  const rows = [];
  presets.forEach(function(preset) {
    Object.keys(preset.values).forEach(function(key) {
      rows.push([preset.preset, key, preset.values[key], notes[key] || '', preset.recommended_for]);
    });
  });

  sheet.clearContents();
  sheet.getRange(1, 1, 1, 5).setValues([['preset', 'key', 'value', 'note', 'recommended_for']]);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 5).setValues(rows);
  }
}

function ensureSheetWithHeaders_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const same = headers.every(function(h, i) {
    return current[i] === h;
  });
  if (!same) {
    sheet.clearContents();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function upsertConfigRows_(sheet, valuesObj, notesObj) {
  const defaults = defaultConfig_();
  const noteMap = notesObj || {};
  const rows = getRows_(sheet);
  const indexByKey = {};
  rows.forEach(function(r, i) {
    const key = String(r.key || '').trim();
    if (key) {
      indexByKey[key] = i;
    }
  });

  const headers = ['key', 'value', 'note'];
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 3).setValues([headers]);
  }

  Object.keys(valuesObj).forEach(function(key) {
    const value = valuesObj[key];
    const note = noteMap[key] || '';
    if (indexByKey[key] != null) {
      const rowNum = indexByKey[key] + 2;
      sheet.getRange(rowNum, 2).setValue(value);
      if (note) {
        sheet.getRange(rowNum, 3).setValue(note);
      }
    } else {
      const rowNum = Math.max(2, sheet.getLastRow() + 1);
      const defaultNote = note || '';
      const fallback = defaults[key];
      sheet.getRange(rowNum, 1, 1, 3).setValues([[key, value != null ? value : fallback, defaultNote]]);
    }
  });
}

function getRows_(sheet) {
  if (!sheet) {
    return [];
  }
  const range = sheet.getDataRange();
  if (range.getNumRows() < 2) {
    return [];
  }
  const values = range.getDisplayValues();
  const headers = values[0];
  return values.slice(1).filter(function(row) {
    return row.join('').trim() !== '';
  }).map(function(row) {
    const obj = {};
    headers.forEach(function(h, i) {
      obj[h] = row[i];
    });
    return obj;
  });
}

function writeRows_(sheet, rows) {
  if (!sheet) {
    return;
  }
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  const headers = headerRange.getValues()[0].filter(function(h) {
    return String(h).trim() !== '';
  });

  sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), Math.max(headers.length, 1)).clearContent();

  if (!rows.length || !headers.length) {
    return;
  }

  const values = rows.map(function(r) {
    return headers.map(function(h) {
      return r[h] == null ? '' : r[h];
    });
  });
  sheet.getRange(2, 1, values.length, headers.length).setValues(values);
}

function writeSummary_(sheet, summaryObj) {
  if (!sheet) {
    return;
  }
  const keys = Object.keys(summaryObj);
  const rows = keys.map(function(k) {
    return [k, summaryObj[k]];
  });

  sheet.clearContents();
  sheet.getRange(1, 1, 1, 2).setValues([['metric', 'value']]);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }
}

function toNumber_(value) {
  if (value == null || value === '') {
    return 0;
  }
  // "¥45", "45円", "1,234.56" などを数値化
  const normalized = String(value)
    .replace(/[,\s]/g, '')
    .replace(/[^\d.\-]/g, '')
    .trim();
  if (normalized === '' || normalized === '.' || normalized === '-' || normalized === '-.') {
    return 0;
  }
  const n = Number(normalized);
  return isNaN(n) ? 0 : n;
}

function toBoolean_(value) {
  const s = String(value || '').trim().toLowerCase();
  return s === '1' || s === 'true' || s === 'yes' || s === 'y';
}

function round2_(num) {
  return Math.round(num * 100) / 100;
}

function calcMigrationBid_(baseBid, roas, config) {
  let bid = toNumber_(baseBid);
  if (bid <= 0) {
    return 0;
  }

  if (config.use_roas_bid_multiplier) {
    const m = getRoasBidMultiplier_(roas, config);
    bid = bid * m;
  }

  if (config.min_bid > 0 && bid < config.min_bid) {
    bid = config.min_bid;
  }
  if (config.max_bid > 0 && bid > config.max_bid) {
    bid = config.max_bid;
  }

  return round2_(bid);
}

function getRoasBidMultiplier_(roas, config) {
  const r = toNumber_(roas);
  if (r >= 35) return config.roas_bid_multiplier_ge_35 || 1;
  if (r >= 25) return config.roas_bid_multiplier_ge_25 || 1;
  if (r >= 20) return config.roas_bid_multiplier_ge_20 || 1;
  if (r >= 15) return config.roas_bid_multiplier_ge_15 || 1;
  return config.roas_bid_multiplier_lt_15 || 1;
}

function normalizeProductKey_(value) {
  return String(value || '').trim().toUpperCase();
}
