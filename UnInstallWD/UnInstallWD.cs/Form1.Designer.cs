namespace UnInstallWD
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
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마십시오.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.label1 = new System.Windows.Forms.Label();
            this.btn_continue = new System.Windows.Forms.Button();
            this.btn_cancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(27, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(275, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "WeDo 프로그램을 제거합니다. 계속하시겠습니까?";
            // 
            // btn_continue
            // 
            this.btn_continue.Location = new System.Drawing.Point(76, 28);
            this.btn_continue.Name = "btn_continue";
            this.btn_continue.Size = new System.Drawing.Size(75, 23);
            this.btn_continue.TabIndex = 1;
            this.btn_continue.Text = "계속";
            this.btn_continue.UseVisualStyleBackColor = true;
            this.btn_continue.MouseClick += new System.Windows.Forms.MouseEventHandler(this.btn_continue_MouseClick);
            // 
            // btn_cancel
            // 
            this.btn_cancel.Location = new System.Drawing.Point(181, 28);
            this.btn_cancel.Name = "btn_cancel";
            this.btn_cancel.Size = new System.Drawing.Size(75, 23);
            this.btn_cancel.TabIndex = 2;
            this.btn_cancel.Text = "취소";
            this.btn_cancel.UseVisualStyleBackColor = true;
            this.btn_cancel.MouseClick += new System.Windows.Forms.MouseEventHandler(this.btn_cancel_MouseClick);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(328, 58);
            this.Controls.Add(this.btn_cancel);
            this.Controls.Add(this.btn_continue);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "WeDo 제거";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btn_continue;
        private System.Windows.Forms.Button btn_cancel;
    }
}

