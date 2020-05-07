namespace QPayBackend.Models
{
    /// <summary>
    /// Confirm class
    /// </summary>
    public class BackendPostData
    {
        /// <summary>
        /// Payment amount
        /// </summary>
        public string ShopNo { get; set; }

        /// <summary>
        /// Payment currency (ISO 4217)
        /// </summary>
        public string PayToken { get; set; }
    }
}