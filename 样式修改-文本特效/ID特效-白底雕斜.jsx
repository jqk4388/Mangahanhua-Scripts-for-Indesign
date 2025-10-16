#include "../Library/KTUlib.jsx"

(function () {
	// 全局常量
	var MM_TO_PT = 1; // 1 mm = 2.834645669 pt

	/**
	 * 主入口
	 */
	function main() {
		try {
			if (app.documents.length === 0) {
				alert('请打开一个文档并选中至少一个文本框。');
				return;
			}

			var doc = app.activeDocument;
			var sel = app.selection;

			if (!sel || sel.length === 0) {
				alert('请先选中一个或多个文本框。');
				return;
			}

			// 遍历选择，针对每个文本框执行操作
			for (var i = 0; i < sel.length; i++) {
				var item = sel[i];
				try {
					processTextFrame(doc, item);
				} catch (e) {
					// 单个对象失败不影响其他对象
					$.writeln('处理对象失败: ' + (e && e.message ? e.message : e));
				}
			}

			// alert('操作完成。');
		} catch (e) {
			alert('脚本发生错误: ' + (e && e.message ? e.message : e));
		}
	}

	/**
	 * 判断对象是否为文本框（TextFrame）
	 * @param {Object} obj - 可能的对象
	 * @returns {Boolean}
	 */
	function isTextFrame(obj) {
		try {
			// 避免使用 instanceof，兼容 ExtendScript
			return obj && obj.hasOwnProperty('contents') && obj.constructor && obj.constructor.name === 'TextFrame';
		} catch (e) {
			return false;
		}
	}

	/**
	 * 对单个文本框执行全部步骤（1-6）
	 * @param {Document} doc - 当前文档
	 * @param {TextFrame} tf - 文本框对象
	 */
	function processTextFrame(doc, tf) {
		try {
			if (!isTextFrame(tf)) {
				// 跳过非文本框
				return;
			}
			var tf1 = tf.duplicate();
			tf.visible = false;
			// 1. 修改字间距为 25（Tracking）并适合框
			applyTrackingAndFit(tf1, 25);

			// 2. 修改文字描边为黑色0.25pt，斜接连接，斜接限制2x，填充白色
			applyTextStrokeAndFill(doc, tf1, 0.25);

			// 3. 多重复制 10 次，每次偏移 -0.02mm, -0.02mm
			var copies = createCopies(tf1, 10, -0.02, -0.02);
			tf1.remove(); // 删除原件
			// 4. 将复制的十个文本框编组，再原地复制一组，记录 id
			var groupA = doc.groups.add(copies);
			var groupB = duplicateInPlace(doc, groupA);
			// 记录 id（ExtendScript 中没有统一 id 属性，但有 id 属性）
			var idA = groupA.id;
			var idB = groupB.id;
			$.writeln('groupA id=' + idA + ' groupB id=' + idB);

			// 5. 将复制的编组移动到后方，并加外发光（应用于 groupB）
			sendToBack(groupB);
			applyOuterGlow(groupB, {
				blendMode: BlendMode.SCREEN, // 滤色
				opacity: 100,
				technique: GlowTechnique.SOFTER, // 柔和
				sizeMM: 0.6,
				noise: 0,
				spread: 100
			});

			// 6. 将十个文本框编组（groupA）和他的外发光编组副本（groupB）一起再编组
            doc.groups.add([groupA, groupB]);

		} catch (e) {
			throw e;
		}
	}

	/**
	 * 复制对象 N 次并按毫米偏移返回复制项数组（不包含原件）
	 * @param {PageItem} item
	 * @param {Number} count - 复制次数
	 * @param {Number} dxMM - 每次复制的 x 偏移（毫米）
	 * @param {Number} dyMM - 每次复制的 y 偏移（毫米）
	 * @returns {Array} 复制得到的 items
	 */
	function createCopies(item, count, dxMM, dyMM) {
		try {
			var results = [];
			var current = item;
			for (var i = 0; i < count; i++) {
				var dx = dxMM * MM_TO_PT;
				var dy = dyMM * MM_TO_PT;
				var dup = current.duplicate();
				// 移动以实现累积偏移
				dup.move(undefined, [dup.geometricBounds[1] + dx - dup.geometricBounds[1], dup.geometricBounds[0] + dy - dup.geometricBounds[0]]);
				results.push(dup);
				current = dup;
			}
			return results;
		} catch (e) {
			throw e;
		}
	}

	/**
	 * 在原地复制一个对象（组）并返回复制对象
	 * @param {Document} doc
	 * @param {PageItem} item
	 * @returns {PageItem}
	 */
	function duplicateInPlace(doc, item) {
		try {
			var dup = item.duplicate();
			// 保持位置不变
			return dup;
		} catch (e) {
			throw e;
		}
	}

	/**
	 * 将页面对象移到后方
	 * @param {PageItem} item
	 */
	function sendToBack(item) {
		try {
			if (!item) return;
			// move to back of its layer
			item.sendToBack();
		} catch (e) {
			throw e;
		}
	}

	/**
	 * 为 group 应用外发光（Outer Glow）效果
	 * @param {PageItem} group
	 * @param {Object} opts - {blendMode, opacity, technique, sizeMM, noise, choke}
	 */
	function applyOuterGlow(group, opts) {
		try {
			if (!group.transparencySettings.outerGlowSettings) {
				// 直接设置属性
			}
			var og = group.transparencySettings.outerGlowSettings;
			og.blendMode = opts.blendMode;
			og.opacity = opts.opacity;
			og.technique = opts.technique;
			og.size = opts.sizeMM * MM_TO_PT;
			og.noise = opts.noise;
			og.spread = opts.spread; 
			// effectColor 设置为白色
			try {
				og.effectColor = group.parent.parent.spreads[0].parentPage.parent.pages[0].parent.parent.documents ? null : og.effectColor;
			} catch (e) {
				// 不强制设置颜色，默认白
			}
			og.applied = true;
		} catch (e) {
			throw e;
		}
	}

	/**
	 * 为文本框中文字应用描边与填充
	 * @param {Document} doc
	 * @param {TextFrame} tf
	 * @param {Number} strokeWeightPt
	 */
	function applyTextStrokeAndFill(doc, tf, strokeWeightPt) {
		try {
			if (!tf || !tf.texts || tf.texts.length === 0) return;
			// 获取全文本范围
			var t = tf.texts[0];
			// 设置填充为白色
			try {
				t.fillColor = doc.swatches.itemByName('Paper');
			} catch (e) {
				// 如果没有 Paper 样本，尝试使用 None -> 使用白色替代
				try {
					var paper = doc.colors.add({name: 'ScriptTempWhite', model: ColorModel.process, colorValue: [0,0,0,0]});
					t.fillColor = paper;
				} catch (e2) {
					// 忽略
				}
			}

			// 设置描边为黑色
			try {
				t.strokeColor = doc.swatches.itemByName('Black');
			} catch (e) {
				// 如果不存在 Black 样本则忽略
			}
			t.strokeWeight = strokeWeightPt;
			// 斜接连接(Miter join)和限制设置
			try {
				t.strokeJoin = StrokeJoin.MITER_JOIN;
				t.miterLimit = 2; // 斜接限制设为 2（2x）
			} catch (e) {
				// 若不支持则忽略
			}

			// Enable stroke on characters
			try {
				t.applyStroke = true;
			} catch (e) {
				// 忽略
			}
		} catch (e) {
			throw e;
		}
	}

	/**
	 * 修改文本框的 Tracking（字距），并如果有 overset（溢流）则 fit
	 * @param {TextFrame} tf
	 * @param {Number} trackingValue - 文字跟踪（追踪），单位 InDesign 跟踪值
	 */
	function applyTrackingAndFit(tf, trackingValue) {
		try {
			if (!tf || !tf.texts || tf.texts.length === 0) return;
			var t = tf.texts[0];
			// 对全文字符设置追踪（tracking）
			try {
				t.tracking = trackingValue;
				//斜变体设置
				t['shataiAdjustRotation'] = false;
				t['shataiAdjustTsume'] = true;
				t['shataiDegreeAngle'] = 4500; //角度45°
				t['shataiMagnification'] = 1500; //放大15%
			} catch (e) {
				// 如果追踪属性不可用，则按字符遍历
				for (var ci = 0; ci < t.characters.length; ci++) {
					try {
						t.characters[ci].tracking = trackingValue;
					} catch (ee) {}
				}
			}

			// 适合文本框：如果溢流则 fit
			try {
				if (tf.overflows) {
					tf.fit(FitOptions.FRAME_TO_CONTENT);
				}
			} catch (e) {
			}
		} catch (e) {
			throw e;
		}
	}

	// 运行主程序
	KTUDoScriptAsUndoable(function() { main(); }, "ID白底雕斜特效");

})();

