using System;
using System.ComponentModel.DataAnnotations;

namespace ImportToDatabase.Models
{
    public class BaseEntity : ICloneable
    {
        public DateTime CreatedDate { get; set; }

        [StringLength(450)]
        public string CreatedBy { get; set; }

        public DateTime? ModifiedDate { get; set; }
        [StringLength(450)]
        public string ModifiedBy { get; set; }

        public object Clone()
        {
            return this.MemberwiseClone();
        }

        public virtual void CalculateUniqKey()
        {

        }
    }
}
