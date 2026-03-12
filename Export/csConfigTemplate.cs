using System;
using System.Collections.Generic;

namespace Logic.TabConfig
{
	//	public class [[tabname]]Mgr : CCfg1KeyMgrTemplate<[[tabname]]Mgr, [[keytype]], [[tabname]]>
	public class SkillConfigInfoMgr : CCfg1KeyMgrTemplate<SkillConfigInfoMgr, int, SkillConfigInfo>
	{
#pragma warning disable
		private void __unused()
		{
			Dictionary<int, SkillConfigInfo> _un = new Dictionary<int, SkillConfigInfo>();
		}
#pragma warning restore
	}

	//	public class [[tabname]] : ITabItemWith1Key<[[keytype]]>
	public class SkillConfigInfo : ITabItemWith1Key<int>
	{
		//public static readonly string __[[列名]] = "[[列名]]";		// [[列的简要注释]]
		public static readonly string __ID = "ID";		// 索引
		public static readonly string __ActionName = "ActionName";
		public static readonly string __CD = "CD";
		public static readonly string __HitID_3_0 = "HitID_3_0";
		public static readonly string __HitTime_3_0 = "HitTime_3_0";
		public static readonly string __HitRangeName_3_0 = "HitRangeName_3_0";
		public static readonly string __HitID_3_1 = "HitID_3_1";
		public static readonly string __HitTime_3_1 = "HitTime_3_1";
		public static readonly string __HitRangeName_3_1 = "HitRangeName_3_1";
		public static readonly string __HitID_3_2 = "HitID_3_2";
		public static readonly string __HitTime_3_2 = "HitTime_3_2";
		public static readonly string __HitRangeName_3_2 = "HitRangeName_3_2";

		//public [[列类型]] [[列名]] { get; private set; }
		public int ID { get; private set; }
		public string ActionName { get; private set; }
		public float CD { get; private set; }
		// 有_x_y的列要转化成数组，x是总数，y是第几个（从0开始—）如HitID_3_0
		//public [[列类型]][] [[列名前缀]] { get; private set; }
		public int[] HitID { get; private set; }
		public float[] HitTime { get; private set; }
		public string[] HitRangeName { get; private set; }

		//	public [[tabname]](){}
		public SkillConfigInfo()
		{
			// 需要初始化数组，如HitID_3_0
			// [[列名前缀]] = new [[列类型]][总数];
			HitID = new int[3];
			HitTime = new float[3];
			HitRangeName = new string[3];
		}

		//	public [[keytype]] GetKey1() { return [[keyname]]; }
		public int GetKey1() { return ID; }

		public bool ReadItem(TabFile tf)
		{
			//	[[列名]] = tf.Get<[[列类型]]>(__[[列名]]);
			ID = tf.Get<int>(__ID);
			ActionName = tf.Get<string>(__ActionName);
			CD = tf.Get<float>(__CD);
			// 注意数组类型的赋值
			//[[列名前缀]][位置] = tf.Get<[[列类型]]>(__[[列名]]);
			HitID[0] = tf.Get<int>(__HitID_3_0);
			HitTime[0] = tf.Get<float>(__HitTime_3_0);
			HitRangeName[0] = tf.Get<string>(__HitRangeName_3_0);
			HitID[1] = tf.Get<int>(__HitID_3_1);
			HitTime[1] = tf.Get<float>(__HitTime_3_1);
			HitRangeName[1] = tf.Get<string>(__HitRangeName_3_1);
			HitID[2] = tf.Get<int>(__HitID_3_2);
			HitTime[2] = tf.Get<float>(__HitTime_3_2);
			HitRangeName[2] = tf.Get<string>(__HitRangeName_3_2);

			return true;
		}
	}
}

