namespace Highlighter
{
    partial class Form1
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.txt_before = new System.Windows.Forms.TextBox();
            this.txt_current = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.pb_loading = new System.Windows.Forms.ProgressBar();
            this.cb_isMerge = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(12, 78);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(82, 31);
            this.button1.TabIndex = 0;
            this.button1.Text = "파일선택";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(12, 150);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(82, 31);
            this.button2.TabIndex = 1;
            this.button2.Text = "파일선택";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // txt_before
            // 
            this.txt_before.Location = new System.Drawing.Point(100, 85);
            this.txt_before.Name = "txt_before";
            this.txt_before.Size = new System.Drawing.Size(276, 21);
            this.txt_before.TabIndex = 2;
            // 
            // txt_current
            // 
            this.txt_current.Location = new System.Drawing.Point(100, 156);
            this.txt_current.Name = "txt_current";
            this.txt_current.Size = new System.Drawing.Size(276, 21);
            this.txt_current.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("굴림", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label1.Location = new System.Drawing.Point(12, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(173, 15);
            this.label1.TabIndex = 4;
            this.label1.Text = "전월/당월 변동사항 확인";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(10, 58);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(119, 12);
            this.label2.TabIndex = 5;
            this.label2.Text = "전월 엑셀파일(.xlsx)";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(14, 130);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(119, 12);
            this.label3.TabIndex = 6;
            this.label3.Text = "당월 엑셀파일(.xlsx)";
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(395, 72);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(70, 103);
            this.button3.TabIndex = 8;
            this.button3.Text = "Start";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // pb_loading
            // 
            this.pb_loading.Location = new System.Drawing.Point(16, 230);
            this.pb_loading.Name = "pb_loading";
            this.pb_loading.Size = new System.Drawing.Size(449, 23);
            this.pb_loading.TabIndex = 9;
            // 
            // cb_isMerge
            // 
            this.cb_isMerge.AutoSize = true;
            this.cb_isMerge.Location = new System.Drawing.Point(17, 197);
            this.cb_isMerge.Name = "cb_isMerge";
            this.cb_isMerge.Size = new System.Drawing.Size(188, 16);
            this.cb_isMerge.TabIndex = 10;
            this.cb_isMerge.Text = "전월엑셀에 당월데이터 합치기";
            this.cb_isMerge.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(485, 271);
            this.Controls.Add(this.cb_isMerge);
            this.Controls.Add(this.pb_loading);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txt_current);
            this.Controls.Add(this.txt_before);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.Text = "전월/당월 변동사항 확인 (문의 : 010-2752-1831)";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox txt_before;
        private System.Windows.Forms.TextBox txt_current;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.ProgressBar pb_loading;
        private System.Windows.Forms.CheckBox cb_isMerge;
    }
}

