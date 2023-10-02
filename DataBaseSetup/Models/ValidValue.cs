namespace Addon_Messem.DataBaseSetup.Models
{
    /// <summary>
    /// Valid Values of a field in SAP B1
    /// </summary>
    public class ValidValue
    {
        /// <summary>
        /// Field Value.
        /// </summary>
        public string Value { get; set; } = string.Empty;

        /// <summary>
        /// The description show to the user.
        /// </summary>
        public string Description { get; set; } = string.Empty;
    }
}
